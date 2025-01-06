// const express = require('express');
// const multer = require('multer');
// const fs = require('fs');
// const csv = require('csv-parser');
// const { v4: uuidv4 } = require('uuid');
// const path = require('path');
// const PizZip = require('pizzip');
// const Docxtemplater = require('docxtemplater');
// const archiver = require('archiver'); // Библиотека для создания ZIP-архивов
// const app = express();
// const upload = multer({ dest: './uploads' });

// app.set('view engine', 'ejs');
// app.set('views', path.join(__dirname, 'views'));

// app.get('/', (req, res) => {
//     res.render('index');
// });

// // Функция для чтения CSV файла
// function readCsvFile(filePath) {
//     return new Promise((resolve, reject) => {
//         const rows = [];
//         fs.createReadStream(filePath, { encoding: 'utf8' })
//             .pipe(csv())
//             .on('data', (data) => {
//                 console.log('CSV Data:', data);
//                 rows.push(data);
//             })
//             .on('end', () => resolve(rows))
//             .on('error', (error) => reject(new Error(`Ошибка чтения CSV файла: ${error.message}`)));
//     });
// }

// // Функция для сопоставления операций
// function matchOperations(requests, registry) {
//     const matched = [];
//     const unmatched = [];
//     requests.forEach((req) => {
//         const match = registry.find((row) => {
//             const cardMatch = row['PAN'] === req['Картка отримувача'] || row['Номер картки'] === req['Картка отримувача'];
//             const tslIdMatch = row['TSL_ID'] === req['TRANID'] || row['Tran_ID'] === req['TRANID'] || row['Унікальнийномер_транзакції_в_ПЦ'] === req['TRANID'];
//             return cardMatch || tslIdMatch;
//         });
//         if (match) {
//             matched.push({
//                 'Дата операції': match['TRAN_DATE_TIME'] || match['Час та датаоперації'],
//                 'Картка отримувача': req['Картка отримувача'],
//                 'Сума зарахування': match['TRAN_AMOUNT'] || match['Сумаоперації,грн.'],
//                 TSL_ID: match['TSL_ID'] || match['Tran_ID'],
//                 'Код авторизації': match['APPROVAL'] || match['Кодавторизації'],
//             });
//         } else {
//             unmatched.push(req);
//         }
//     });
//     return { matched, unmatched };
// }

// // Функция для генерации документов
// function generateDocument(matchedData) {
//     return new Promise((resolve, reject) => {
//         const templatePath = path.join(__dirname, 'template.docx');
//         const templateBuffer = fs.readFileSync(templatePath);
//         try {
//             const zip = new PizZip(templateBuffer);
//             const generatedFiles = [];

//             matchedData.forEach((data, index) => {
//                 const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
//                 try {
//                     doc.setData({
//                         //                         'DataOperatsiyi': data['Дата операції'],
//                         //                         'KartkaOtrymuvacha': data['Картка отримувача'],
//                         //                         'SumaZarakhuvannya': data['Сума зарахування'],
//                         //                         'TSL_ID': data['TSL_ID'],
//                         //                         'KodAvtoryzatsiyi': data['Код авторизації'],


//                         "DataOperatsiyi": "512647352",
//                         "KartkaOtrymuvacha": "ТОВ «1 БЕЗПЕЧНЕ АГЕНТСТВО НЕОБХІДНИХ КРЕДИТІВ»",
//                         "SumaZarakhuvannya": "12.07.2024 20:36",
//                         "TSL_ID": "6000",
//                         "KodAvtoryzatsiyi": "432335****** 9523",
//                         "company":"1 БЕЗПЕЧНЕ АГЕНТСТВО НЕОБХІДНИХ КРЕДИТІВ",
//                         "sender_list":"№05/11/24-18 від 05.11.2024",
//                         "document":"про організацію взаємодії при переказі коштів фізичним особам №74 від «03» серпня 2020р",
//                         "additinal":"ЗВАРИЧУК БОГДАН АНДРІЙОВИЧ;  код  8a27546a-5b64-4cef-8c59-669169ba75a3, сума 14500,00 грн., дата транзакції 12.07.2024, договір №79176882, дата договору 12.07.2024, маскований номер картки 432335****** 9523.",
//                         "sender": "Чугунов Василь Васильович"
//                     });

//                     doc.render();
//                     const buf = doc.getZip().generate({ type: 'nodebuffer' });
//                     const fileName = `output_${index + 1}.docx`;
//                     generatedFiles.push({ fileName, buffer: buf });
//                 } catch (error) {
//                     reject(error); 
//                 }
//             });

//             resolve(generatedFiles); 
//         } catch (error) {
//             reject(error);
//         }
//     });
// }

// // Маршрут для загрузки файлов
// app.post('/upload', upload.fields([
//     { name: 'registryFiles', maxCount: 1 },
//     { name: 'requestFile', maxCount: 1 }
// ]), async (req, res) => {
//     try {
//         console.log(req.files);
//         if (!req.files['registryFiles'] || !req.files['requestFile']) {
//             throw new Error('Не все файлы были загружены');
//         }
//         const registryFile = req.files['registryFiles'][0];
//         const requestFile = req.files['requestFile'][0];

//         // Чтение данных из CSV файлов
//         const registryData = await readCsvFile(registryFile.path);
//         const requestData = await readCsvFile(requestFile.path);

//         // Сопоставление операций
//         const { matched, unmatched } = matchOperations(requestData, registryData);

//         // Генерация документов для каждой совпавшей операции
//         const generatedFiles = await generateDocument(matched);

//         // Создание архива
//         const archive = archiver('zip', { zlib: { level: 9 } });
//         res.setHeader('Content-Type', 'application/zip');
//         res.setHeader('Content-Disposition', 'attachment; filename=documents.zip');
//         archive.pipe(res);

//         // Добавление файлов в архив
//         generatedFiles.forEach(file => {
//             archive.append(file.buffer, { name: file.fileName });
//         });

//         archive.finalize();
//     } catch (error) {
//         console.error('Ошибка обработки файлов:', error);
//         res.status(500).json({ error: error.message });
//     }
// });

// const PORT = 3000;
// app.listen(PORT, () => {
//     console.log(`Сервер запущен на порту ${PORT}`);
// });


// const express = require('express');
// const multer = require('multer');
// const fs = require('fs');
// const csv = require('csv-parser');
// const { v4: uuidv4 } = require('uuid');
// const path = require('path');
// const PizZip = require('pizzip');
// const Docxtemplater = require('docxtemplater');
// const archiver = require('archiver');
// const app = express();
// const upload = multer({ dest: './uploads' });

// app.set('view engine', 'ejs');
// app.set('views', path.join(__dirname, 'views'));

// app.get('/', (req, res) => {
//     res.render('index');
// });

// // Функция для чтения CSV файла
// function readCsvFile(filePath) {
//     return new Promise((resolve, reject) => {
//         const rows = [];
//         fs.createReadStream(filePath, { encoding: 'utf8' })
//             .pipe(csv())
//             .on('data', (data) => {
//                 rows.push(data);
//             })
//             .on('end', () => resolve(rows))
//             .on('error', (error) => reject(new Error(`Ошибка чтения CSV файла: ${error.message}`)));
//     });
// }

// function matchOperations(requests, registry) {
//     const matched = [];
//     const unmatched = [];
//     requests.forEach((req) => {
//         const match = registry.find((row) => {
//             const cardMatch = row['PAN'] === req['Картка отримувача'] || row['Номер картки'] === req['Картка отримувача'];
//             const tslIdMatch = row['TSL_ID'] === req['TRANID'] || row['Tran_ID'] === req['TRANID'] || row['Унікальнийномер_транзакції_в_ПЦ'] === req['TRANID'];
//             return cardMatch || tslIdMatch;
//         });
//         if (match) {
//             matched.push({
//                 'Дата операції': match['TRAN_DATE_TIME'] || match['Час та датаоперації'],
//                 'Картка отримувача': req['Картка отримувача'],
//                 'Сума зарахування': match['TRAN_AMOUNT'] || match['Сумаоперації,грн.'],
//                 TSL_ID: match['TSL_ID'] || match['Tran_ID'],
//                 'Код авторизації': match['APPROVAL'] || match['Кодавторизації'],
//                 company: req['Назва компанії'],
//                 sender_list: req['Номер листа'],
//                 document: req['Додаткова інформація'],
//                 additinal: req['Додаткова інформація'],
//                 sender: req['ПІБ']
//             });
//         } else {
//             unmatched.push(req);
//         }
//     });
//     console.log("Matched:", matched);  // Логируем совпавшие данные
//     console.log("Unmatched:", unmatched);  // Логируем не совпавшие данные
//     return { matched, unmatched };
// }

// function generateDocument(matchedData) {
//     return new Promise((resolve, reject) => {
//         const templatePath = path.join(__dirname, 'template.docx');
//         const templateBuffer = fs.readFileSync(templatePath);
//         try {
//             const zip = new PizZip(templateBuffer);
//             const generatedFiles = [];

//             matchedData.forEach((data, index) => {
//                 const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
//                 try {
//                     // Выводим все данные в консоль, чтобы проверить их
//                     console.log("Data to be set into document:", data);

//                     const dataToSet = {
//                         "DataOperatsiyi": data['Дата операції'] || "Default value",
//                         "KartkaOtrymuvacha": data['Картка отримувача'] || "Default value",
//                         "SumaZarakhuvannya": data['Сума зарахування'] || "Default value",
//                         "TSL_ID": data['TSL_ID'] || "Default value",
//                         "KodAvtoryzatsiyi": data['Код авторизації'] || "Default value",
//                         "company": data['Назва компанії'] || "Default value",
//                         "sender_list": data['Номер листа'] || "Default value",
//                         "document": data['Додаткова інформація'] || "Default value",
//                         "additinal": data['Додаткова інформація'] || "Default value",
//                         "sender": data['ПІБ'] || "Default value"
//                     };
//                     console.log("Data to be set:", dataToSet);
//                     doc.render(dataToSet);  // Вызываем render один раз

//                     const buf = doc.getZip().generate({ type: 'nodebuffer' });
//                     const fileName = `output_${index + 1}.docx`;
//                     generatedFiles.push({ fileName, buffer: buf });
//                 } catch (error) {
//                     reject(error);
//                 }
//             });

//             resolve(generatedFiles); 
//         } catch (error) {
//             reject(error);
//         }
//     });
// }



// // Маршрут для загрузки файлов
// app.post('/upload', upload.fields([
//     { name: 'registryFiles', maxCount: 1 },
//     { name: 'requestFile', maxCount: 1 }
// ]), async (req, res) => {
//     try {
//         if (!req.files['registryFiles'] || !req.files['requestFile']) {
//             throw new Error('Не все файлы были загружены');
//         }
//         const registryFile = req.files['registryFiles'][0];
//         const requestFile = req.files['requestFile'][0];

//         // Чтение данных из CSV файлов
//         const registryData = await readCsvFile(registryFile.path);
//         const requestData = await readCsvFile(requestFile.path);

//         // Сопоставление операций
//         const { matched, unmatched } = matchOperations(requestData, registryData);

//         // Генерация документов для каждой совпавшей операции
//         const generatedFiles = await generateDocument(matched);

//         // Создание архива
//         const archive = archiver('zip', { zlib: { level: 9 } });
//         res.setHeader('Content-Type', 'application/zip');
//         res.setHeader('Content-Disposition', 'attachment; filename=documents.zip');
//         archive.pipe(res);

//         // Добавление файлов в архив
//         generatedFiles.forEach(file => {
//             archive.append(file.buffer, { name: file.fileName });
//         });

//         archive.finalize();
//     } catch (error) {
//         console.error('Ошибка обработки файлов:', error);
//         res.status(500).json({ error: error.message });
//     }
// });

// const PORT = 3000;
// app.listen(PORT, () => {
//     console.log(`Сервер запущен на порту ${PORT}`);
// });


const express = require('express');
const multer = require('multer');
const fs = require('fs');
const csv = require('csv-parser');
const { v4: uuidv4 } = require('uuid');
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

// Функция для чтения CSV файла
function readCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath, { encoding: 'utf8' })
            .pipe(csv())  // Используем csv-parser для чтения
            .on('data', (data) => {
                rows.push(data);
            })
            .on('end', () => {
                console.log("CSV Data read from file:", rows);  // Логируем данные, считанные из CSV
                resolve(rows);
            })
            .on('error', (error) => reject(new Error(`Ошибка чтения CSV файла: ${error.message}`)));
    });
}

// Сопоставление TRANID из запроса с данными реестра
function matchOperations(requests, registry) {
    const matched = [];
    const unmatched = [];

    requests.forEach((req) => {
        const transactionId = req['TRANID'];  // Берем TRANID из запроса
        const match = registry.find((row) => {
            // Проверяем совпадения по разным возможным столбцам реестра
            const tslIdMatch = (row['TSL_ID'] || '').trim() === (transactionId || '').trim();
            const tranIdMatch = (row['Tran_ID'] || '').trim() === (transactionId || '').trim();
            const uniqueTransactionIdMatch = (row['Унікальнийномер_транзакції_в_ПЦ'] || '').trim() === (transactionId || '').trim();
            
            return tslIdMatch || tranIdMatch || uniqueTransactionIdMatch;
        });

        if (match) {
            // Если нашли совпадение, добавляем информацию в массив matched
            matched.push({
                'Дата операції': match['TRAN_DATE_TIME'] || match['Час та датаоперації'],
                'Картка отримувача': req['Картка отримувача'],
                'Сума зарахування': match['TRAN_AMOUNT'] || match['Сумаоперації,грн.'],
                TSL_ID: match['TSL_ID'] || match['Tran_ID'],
                'Код авторизації': match['APPROVAL'] || match['Кодавторизації'],
                company: req['Назва компанії'],
                sender_list: req['Номер листа'],
                document: req['Додаткова інформація'],
                additional: req['Додаткова інформація'],
                sender: req['ПІБ']
            });
        } else {
            // Если не нашли совпадение, добавляем в unmatched
            unmatched.push(req);
        }
    });

    console.log("Matched:", matched);  // Логируем совпавшие данные
    console.log("Unmatched:", unmatched);  // Логируем не совпавшие данные
    return { matched, unmatched };
}

function generateDocument(matchedData) {
    return new Promise((resolve, reject) => {
        const templatePath = path.join(__dirname, 'template.docx');
        const templateBuffer = fs.readFileSync(templatePath);
        try {
            const zip = new PizZip(templateBuffer);
            const generatedFiles = [];

            matchedData.forEach((data, index) => {
                const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
                try {
                    // Выводим все данные в консоль, чтобы проверить их
                    console.log("Data to be set into document:", data);

                    const dataToSet = {
                        "DataOperatsiyi": data['Дата операції'] || "Default value",
                        "KartkaOtrymuvacha": data['Картка отримувача'] || "Default value",
                        "SumaZarakhuvannya": data['Сума зарахування'] || "Default value",
                        "TSL_ID": data['TSL_ID'] || "Default value",
                        "KodAvtoryzatsiyi": data['Код авторизації'] || "Default value",
                        "company": data['Назва компанії'] || "Default value",
                        "sender_list": data['Номер листа'] || "Default value",
                        "document": data['Додаткова інформація'] || "Default value",
                        "additinal": data['Додаткова інформація'] || "Default value",
                        "sender": data['ПІБ'] || "Default value"
                    };
                    console.log("Data to be set into document:", dataToSet); // Логируем данные, которые будем вставлять в документ

                    doc.render(dataToSet);  // Вызываем render один раз

                    const buf = doc.getZip().generate({ type: 'nodebuffer' });
                    const fileName = `output_${index + 1}.docx`;
                    generatedFiles.push({ fileName, buffer: buf });

                    console.log(`Generated document ${fileName}`);  // Логируем файл, который был сгенерирован

                } catch (error) {
                    reject(error);
                }
            });

            resolve(generatedFiles); 
        } catch (error) {
            reject(error);
        }
    });
}


// Маршрут для загрузки файлов
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

        // Чтение данных из CSV файлов
        const registryData = await readCsvFile(registryFile.path);
        const requestData = await readCsvFile(requestFile.path);

        // Сопоставление операций
        const { matched, unmatched } = matchOperations(requestData, registryData);

        // Генерация документов для каждой совпавшей операции
        const generatedFiles = await generateDocument(matched);

        // Создание архива
        const archive = archiver('zip', { zlib: { level: 9 } });
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename=documents.zip');
        archive.pipe(res);

        // Добавление файлов в архив
        generatedFiles.forEach(file => {
            archive.append(file.buffer, { name: file.fileName });
            console.log(`Adding file ${file.fileName} to the archive`);  // Логируем добавление файла в архив
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


