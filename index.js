const express = require('express');
const multer = require('multer');
const fs = require('fs');
const csv = require('csv-parser');
const path = require('path');
const PizZip = require('pizzip');
const iconv = require('iconv-lite');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');
const app = express();
const upload = multer({ dest: './uploads' });
app.use(express.static(path.join(__dirname, 'public')));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.get('/processing/', (req, res) => {
    res.render('index');
});

function readCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        // Чтение файла с конвертацией из кодировки Windows-1251 в UTF-8
        const readStream = fs.createReadStream(filePath)
            .pipe(iconv.decodeStream('win1251'))  // Декодируем поток из Windows-1251
            .pipe(iconv.encodeStream('utf8'));   // Перекодируем в UTF-8

        readStream
            .pipe(csv({ separator: ';' })) // Разделитель для CSV
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

                // const processedData = {
                //     'Дата операції': match['TRAN_DATE_TIME'] || match['Час та датаоперації'],
                //     'Картка отримувача': match['PAN'] || match['Номер картки'],
                //     'Сума зарахування': match['TRAN_AMOUNT'] || match['Сумаоперації,грн.'],
                //     // TSL_ID: match['TSL_ID'] || match['TRANID'] || match['Уникальныйномертранзакции в ПЦ'],
                //     TSL_ID: match['TSL_ID'] || match['Уникальныйномертранзакции в ПЦ'] || match['TRANID'],

                //     'Код авторизації': match['APPROVAL'] || match['Кодавторизації'],
                //     company: fullCompanyName, 
                //     companyShort: fullCompanyName.slice(0, -16).trim() ,
                //     sender_list: request['Номер вихідного листа'],
                //     document: request['Додаткова інформація'],
                //     additional: request['Додаткова інформація'],
                //     sender: request['ПІБ клієнта'],
                //     name: request['ПІБ клоієнта'],
                //     dogovir: request['Договір']
                // };

                const processedData = {
                    'Дата операції': match['TRAN_DATE_TIME'] || match['Час та датаоперації'],
                    'Картка отримувача': match['PAN'] || match['Номер картки'],
                    'Сума зарахування': match['TRAN_AMOUNT'] || match['Сумаоперації,грн.'],
                    TSL_ID: match['TSL_ID'] || match['Уникальныйномертранзакции в ПЦ'] || match['TRANID'],
                    'Код авторизації': match['APPROVAL'] || match['Кодавторизації'],
                    company: fullCompanyName, 
                    companyShort: fullCompanyName.slice(0, -16).trim(),
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
    { name: 'registryFiles', maxCount: 10 },
    { name: 'requestFile', maxCount: 10 }
]), async (req, res) => {
    try {
        // Проверяем наличие загруженных файлов
        if (!req.files['registryFiles'] || !req.files['requestFile']) {
            return res.status(400).json({ error: 'Не все файлы были загружены' });
        }

        const registryFiles = req.files['registryFiles'];
        const requestFiles = req.files['requestFile'];

        // Читаем данные из файлов
        const registryDataPromises = registryFiles.map(file => readCsvFile(file.path));
        const requestDataPromises = requestFiles.map(file => readCsvFile(file.path));

        const [registryDataArray, requestDataArray] = await Promise.all([
            Promise.all(registryDataPromises),
            Promise.all(requestDataPromises)
        ]);

        const matched = [];
        const unmatched = [];

        // Сопоставляем данные реестра и запросов
        registryDataArray.forEach(registryData => {
            requestDataArray.forEach(requestData => {
                const matchResult = matchOperations(requestData, registryData);
                matched.push(...matchResult.matched);
                unmatched.push(...matchResult.unmatched);
            });
        });


        if (req.query.report === 'true') {
            const report = {
                matched: matched.map(item => ({
                    name: item.sender || "Не указано",
                    status: "Найдено"
                })),
                unmatched: unmatched
                    .filter(item => item['ПІБ клієнта'] && item['ПІБ клієнта'] !== "") 
                    .map(item => ({
                        name: item['ПІБ клієнта'],
                        status: "Не найдено"
                    }))
            };
        
            return res.status(200).json(report);
        }
        
        



        // Если совпадений нет, возвращаем соответствующий ответ
        if (matched.length === 0) {
            return res.status(404).json({ message: 'Совпадений не найдено' });
        }

        // Определяем шаблон документа в зависимости от типа
        const templatePath = req.query.type === 'c2a'
            ? path.join(__dirname, 'template_c2a.docx')
            : req.query.type === 'a2c'
                ? path.join(__dirname, 'template_a2c.docx')
                : null;

        if (!templatePath) {
            return res.status(400).json({ error: 'Неверный тип запроса. Поддерживаются только c2a и a2c' });
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
                    "sender": data.sender || "Default value",
                    "companyShort": data.companyShort || "Default value",
                    "dogovir": data.dogovir || "Default value",
                });

                const buffer = doc.getZip().generate({ type: 'nodebuffer' });
                const fileName = `${data.sender || 'default_name'}.docx`;

                generatedFiles.push({ fileName, buffer });
            });

            return generatedFiles;
        };

        const generatedFiles = await generateDocumentWithTemplate(matched, templatePath);

        // Создание ZIP-архива с файлами
        const archive = archiver('zip', { zlib: { level: 9 } });
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename=documents.zip');
        archive.pipe(res);

        generatedFiles.forEach(file => {
            archive.append(file.buffer, { name: file.fileName });
        });

        archive.finalize();
    } catch (error) {
        console.error('Ошибка обработки файлов:', error);
        res.status(500).json({ error: error.message });
    }
});



const PORT = 1515;
app.listen(PORT, () => {
    console.log(`Сервер запущен на порту ${PORT}`);
});
