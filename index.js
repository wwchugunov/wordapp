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
                matched.push({
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
                });
            }

        } else {
            unmatched.push(request);
        }
    });
    return { matched, unmatched };
}

function generateDocument(matchedData) {
    return new Promise((resolve, reject) => {
        const templatePath = path.join(__dirname, 'template.docx');
        if (!fs.existsSync(templatePath)) {
            return reject(new Error('Шаблон документа не найден'));
        }

        const templateBuffer = fs.readFileSync(templatePath);
        try {
            const generatedFiles = [];
            const processedTSLIds = new Set(); 

            matchedData.forEach((data) => {
                if (!Object.values(data).some(value => value) || !data['Сума зарахування'] || !data['TSL_ID']) {
                    console.log(`Пропуск пустого документа для TRANID: ${data.TSL_ID || 'неизвестно'}`);
                    return;
                }

                if (processedTSLIds.has(data.TSL_ID)) {
                    console.log(`Документ для TSL_ID ${data.TSL_ID} уже был сгенерирован. Пропуск.`);
                    return;
                }
                processedTSLIds.add(data.TSL_ID); 

                const clientName = data.sender || 'default_name'; 

                const fileName = `${clientName}.docx`;

                try {

                    const zip = new PizZip(templateBuffer);
                    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

                    console.log("Данные для вставки в документ:", data);

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

                    const buf = doc.getZip().generate({ type: 'nodebuffer' });
                    generatedFiles.push({ fileName, buffer: buf });

                    console.log(`Сгенерирован документ: ${fileName}`);

                } catch (error) {
                    console.error(`Ошибка генерации документа для TRANID: ${data.TSL_ID}`, error);
                }
            });

            resolve(generatedFiles);
        } catch (error) {
            reject(error);
        }
    });
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

        const generatedFiles = await generateDocument(matched);

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
