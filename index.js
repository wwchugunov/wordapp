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

// Функция для чтения CSV файла
function readCsvFile(filePath) {
    return new Promise((resolve, reject) => {
        const rows = [];
        fs.createReadStream(filePath, { encoding: 'utf8' })
            .pipe(csv({ separator: ';' })) // Указываем разделитель
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
        const tranId = (request['TRANID'] || '').trim(); // Убираем пробелы
        const match = registry.find(row => {
            const tslId = (row['TSL_ID'] || '').trim(); // Убираем пробелы
            const regTranId = (row['TRANID'] || '').trim(); // Убираем пробелы

            return tranId === tslId || tranId === regTranId; // Сравниваем по TRANID
        });

        if (match) {
            // Проверяем, чтобы одно совпадение для TRANID не повторялось
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

const sanitizeFileName = (name) => {
    // Заменяем пробелы и специальные символы на подчеркивания или убираем их
    return name.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 100); // Ограничиваем длину до 100 символов
}

function generateDocument(matchedData) {
    return new Promise((resolve, reject) => {
        const templatePath = path.join(__dirname, 'template.docx');
        if (!fs.existsSync(templatePath)) {
            return reject(new Error('Шаблон документа не найден'));
        }

        const templateBuffer = fs.readFileSync(templatePath);
        try {
            const zip = new PizZip(templateBuffer);
            const generatedFiles = [];
            const processedTSLIds = new Set(); // Множество для отслеживания уникальных TSL_ID

            matchedData.forEach((data) => {
                // Пропускаем данные с отсутствующими или пустыми значениями
                if (!Object.values(data).some(value => value) || !data['Сума зарахування'] || !data['TSL_ID']) {
                    console.log(`Пропуск пустого документа для TRANID: ${data.TSL_ID || 'неизвестно'}`);
                    return;
                }

                // Проверяем уникальность TSL_ID
                if (processedTSLIds.has(data.TSL_ID)) {
                    console.log(`Документ для TSL_ID ${data.TSL_ID} уже был сгенерирован. Пропуск.`);
                    return;
                }
                processedTSLIds.add(data.TSL_ID); // Добавляем TSL_ID в множество

                // Используем поле "ПІБ клієнта" как имя файла
                const clientName = data['ПІБ клієнта'] || 'default_name'; // Если "ПІБ клієнта" отсутствует, используем значение по умолчанию
                if (!clientName.trim()) {
                    console.log('Ошибка: ПІБ клієнта пустое или отсутствует');
                    return;
                }
                const sanitizedClientName = sanitizeFileName(clientName); // Очистка имени

                const fileName = `${sanitizedClientName}.docx`; // Уникальное имя файла, основанное на "ПІБ клієнта"
                const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

                try {
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
