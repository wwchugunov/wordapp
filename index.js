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
                const processedData = {
                    'Дата операції': match['STL_DATE'] || match['Час операції'],
                    'Картка отримувача': match['PAN'],
                    'Сума зарахування': match['TRAN_AMOUNT'] || match['Сума'],
                    TSL_ID: match['TSL_ID'] || match['TRANID'],
                    'Код авторизації': match['APPROVAL'] || '',
                    company: Object.keys(request).find(key => /Юридична\s+назва\s+ЄДРПОУ/i.test(key))
                        ? request[Object.keys(request).find(key => /Юридична\s+назва\s+ЄДРПОУ/i.test(key))]
                        : 'Default value',
                    sender_list: request['Номер вихідного листа'],
                    document: request['Додаткова інформація'],
                    additional: request['Додаткова інформація'],
                    sender: request['ПІБ клієнта'],
                    name: request['ПІБ клієнта']
                };

                if (
                    processedData['Сума зарахування'] &&
                    processedData.TSL_ID &&
                    processedData.sender
                ) {
                    matched.push(processedData);
                } else {
                    console.log(`Пропуск некорректной записи: ${JSON.stringify(processedData)}`);
                }
            }
        } else {
            unmatched.push(request);
        }
    });

    return { matched, unmatched };
}


app.post('/upload', upload.fields([
    { name: 'registryFiles', maxCount: 1 },
    { name: 'requestFile', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files['registryFiles'] || !req.files['requestFile']) {
            throw new Error('Не все файлы были загружены');
        }

        const registryFile = req.files['registryFiles'][0];
        const requestFile = req.files['requestFile'][0];

        const registryData = await readCsvFile(registryFile.path);
        const requestData = await readCsvFile(requestFile.path);

        const { matched, unmatched } = matchOperations(requestData, registryData);

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
                    "sender": data.sender || "Default value"
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
