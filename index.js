const express = require('express');
const multer = require('multer');
const fs = require('fs');
const csv = require('csv-parser');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');
const app = express();
const upload = multer({ dest: './uploads' });
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.get('/', (req, res) => {
    res.render('index');
});

function readCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath, { encoding: 'utf8' })
            .pipe(csv({ separator: ';' })) 
            .on('data', (data) => {
                rows.push(data);
            })
            .on('end', () => {
                console.log("CSV Data read from file:", rows);
                resolve(rows);
            })
            .on('error', (error) => reject(new Error(`Ошибка чтения CSV файла: ${error.message}`)));
    });
}



function matchOperations(requests, registry) {
    const matched = [];
    const unmatched = [];

    requests.forEach(request => {
        const tranId = (request['TRANID'] || '').trim(); 
        const match = registry.find(row => {
            const tslId = (row['TSL_ID'] || '').trim(); 
            const regTranId = (row['TRANID'] || '').trim(); 

            return tranId === tslId || tranId === regTranId;
        });

        if (match) {
            const isAlreadyMatched = matched.some(item => item.TSL_ID === match['TSL_ID']);

            if (!isAlreadyMatched) {
                const fullCompanyName = Object.keys(request).find(key => /Юридична\s+назва\s+ЄДРПОУ/i.test(key))
                    ? request[Object.keys(request).find(key => /Юридична\s+назва\s+ЄДРПОУ/i.test(key))]
                    : 'Default value';

                const processedData = {
                    'Дата операції': match['TRAN_DATE_TIME'] || match['Час операції'],
                    'Картка отримувача': match['PAN'],
                    'Сума зарахування': match['TRAN_AMOUNT'] || match['Сума'],
                    TSL_ID: match['TSL_ID'] || match['TRANID'],
                    'Код авторизації': match['APPROVAL'] || '',
                    company: fullCompanyName, // Полное название компании
                    companyShort: fullCompanyName.slice(0, -16).trim() , // Последние 8 символов названия компании
                    sender_list: request['Номер вихідного листа'],
                    document: request['Додаткова інформація'],
                    additional: request['Додаткова інформація'],
                    sender: request['ПІБ клієнта'],
                    name: request['ПІБ клоієнта'],
                    dogovir: request['Договір']
                };

                console.log(processedData)
                if (
                    processedData['Сума зарахування'] &&
                    processedData.sender
                ) {
                    matched.push(processedData);
                } else {
                    console.log(processedData);
                }
            }

        } else {
            unmatched.push(request);
        }
    });

    return { matched, unmatched };
}




app.post('/upload', upload.fields([
    { name: 'registryFiles', maxCount: 10 },  // Задаем лимит на 10 файлов
    { name: 'requestFile', maxCount: 10 }     // Задаем лимит на 10 файлов
]), async (req, res) => {
    try {
        if (!req.files['registryFiles'] || !req.files['requestFile']) {
            throw new Error('Не все файлы были загружены');
        }

        // Извлекаем все файлы реестра и запросов партнера
        const registryFiles = req.files['registryFiles'];
        const requestFiles = req.files['requestFile'];

        // Прочитаем все файлы по очереди
        const registryDataPromises = registryFiles.map(file => readCsvFile(file.path));
        const requestDataPromises = requestFiles.map(file => readCsvFile(file.path));
        
        const registryDataArray = await Promise.all(registryDataPromises);
        const requestDataArray = await Promise.all(requestDataPromises);

        // Теперь у нас есть все данные из файлов реестра и запросов
        const matched = [];
        const unmatched = [];

        // Пройдем по всем парам данных реестра и запроса
        registryDataArray.forEach(registryData => {
            requestDataArray.forEach(requestData => {
                const matchResult = matchOperations(requestData, registryData);
                matched.push(...matchResult.matched);
                unmatched.push(...matchResult.unmatched);
            });
        });

        if (matched.length === 0) {
            return res.status(404).json({ message: 'Совпадений не найдено' });
        }

        let templatePath;
        if (req.query.type === 'c2a') {
            templatePath = path.join(__dirname, 'template_c2a.docx');
        } else if (req.query.type === 'a2c') {
            templatePath = path.join(__dirname, 'template_a2c.docx');
        } else {
            throw new Error('Неверный тип запроса. Поддерживаются только c2a и a2c');
        }

        const generateDocumentWithTemplate = async (matchedData, templatePath) => {
            const templateBuffer = fs.readFileSync(templatePath);

            const generatedFiles = [];
            const processedTSLIds = new Set();

            matchedData.forEach((data) => {
                if (!data['Сума зарахування'] || processedTSLIds.has(data.TSL_ID)) return;

                processedTSLIds.add(data.TSL_ID);
                const zip = new PizZip(templateBuffer);
                const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

                // doc.render({
                //     "DataOperatsiyi": data['Дата операції'] || "Default value",
                //     "KartkaOtrymuvacha": data['Картка отримувача'] || "Default value",
                //     "SumaZarakhuvannya": data['Сума зарахування'] || "Default value",
                //     "TSL_ID": data['TSL_ID'] || "Default value",
                //     "KodAvtoryzatsiyi": data['Код авторизації'] || "Default value",
                //     "company": data.company || "Default value",
                //     "sender_list": data.sender_list || "Default value",
                //     "document": data.document || "Default value",
                //     "additional": data.additional || "Default value",
                //     "sender": data.sender || "Default value",
                //     "companyShort": data.companyShort || "Default value",
                //     "dogovir": data.dogovir || "Default value",
                // });

                doc.render({
                    "DataOperatsiyi": data['Дата операції'] || "Default value",
                    "KartkaOtrymuvacha": data['Картка отримувача'] || "Default value",
                    "SumaZarakhuvannya": data['Сума зарахування'] || "Default value",
                    "TSL_ID": data['TSL_ID'] || "Default value",
                    "KodAvtoryzatsiyi": data['Код авторизації'] || "Default value",
                    "company": data.company || "Default value",
                    "sender_list": data.sender_list || "Default value",
                    "document": data.document || "Default value",
                    "additional": data.additional || "Default value",
                    "sender": data.sender || "Default value",
                    "companyShort": data.companyShort || "Default value", // Здесь передаётся обрезанное название
                    "dogovir": data.dogovir || "Default value",
                });

                console.log({
                    company: data.company,
                    companyShort: data.companyShort,
                    dogovir: data.dogovir,
                });
                
                

                const buffer = doc.getZip().generate({ type: 'nodebuffer' });
                const fileName = `${data.sender || 'default_name'}.docx`;

                generatedFiles.push({ fileName, buffer });
            });

            return generatedFiles;
        };

        const generatedFiles = await generateDocumentWithTemplate(matched, templatePath);

        const archive = archiver('zip', { zlib: { level: 9 } });
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename=documents.zip');
        archive.pipe(res);

        generatedFiles.forEach(file => {
            archive.append(file.buffer, { name: file.fileName });
            console.log(`Добавлен файл ${file.fileName} в архив`);
        });

        archive.finalize();
    } catch (error) {
        console.error('Ошибка обработки файлов:', error);
        res.status(500).json({ error: error.message });
    }
});



const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Сервер запущен на порту ${PORT}`);
});
