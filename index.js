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

app.get('/formater', (req, res) => {
    res.render('index');
});

function readCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath, { encoding: 'utf8' })
            .pipe(csv())  
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

    requests.forEach((req) => {
        const transactionId = req['TRANID'];  
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
                    console.log("Data to be set into document:", dataToSet); 


                    const buf = doc.getZip().generate({ type: 'nodebuffer' });
                    const fileName = `output_${index + 1}.docx`;
                    generatedFiles.push({ fileName, buffer: buf });


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

const PORT = 3060;
app.listen(PORT, () => {
    console.log(`Сервер запущен на порту ${PORT}`);
});


