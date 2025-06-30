const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const os = require('os');

// Функция для получения доступного для записи пути
function getWritablePath() {
    // Пробуем несколько возможных путей
    const possiblePaths = [
        process.cwd() // Текущая рабочая директория
    ];

    for (const dir of possiblePaths) {
        try {
            // Проверяем, доступен ли путь для записи, создавая пустой временный файл
            const testFile = path.join(dir, `.test-${Date.now()}.tmp`);
            fs.writeFileSync(testFile, '');
            fs.unlinkSync(testFile); // Удаляем тестовый файл
            return dir; // Если успешно, возвращаем этот путь
        } catch (error) {
            console.log(`Путь ${dir} недоступен для записи: ${error.message}`);
            // Продолжаем с следующим путем
        }
    }

    throw new Error('Не найдено доступного пути для записи файлов');
}

// Определяем возможные пути для поиска Excel файлов
function getSourceDirectories() {
    const dirs = [];

    // Путь относительно исполняемого файла
    try {
        const exePath = process.execPath;
        const exeDir = path.dirname(exePath);
        dirs.push(path.join(exeDir, 'Source'));
    } catch (error) {
        console.log('Не удалось определить путь исполняемого файла');
    }

    // Проверяем, какие директории существуют
    return dirs.filter(dir => {
        try {
            return fs.existsSync(dir);
        } catch (error) {
            return false;
        }
    });
}

// Source directory where Excel files are located
// Будем искать в нескольких местах
const sourceDirs = getSourceDirectories();

// Output file path с временной меткой для уникальности
const timestamp = new Date().getTime();
const outputDir = getWritablePath();
const outputFile = path.join(outputDir, `merged_data_${timestamp}.xlsx`);

// Function to read all Excel files from available directories
function readExcelFiles() {
    if (sourceDirs.length === 0) {
        console.log('Папка Source не найдена. Создаем её...');
        try {
            fs.mkdirSync(path.join(process.cwd(), 'Source'));
            console.log(`Создана папка Source в ${process.cwd()}`);
            console.log('Пожалуйста, поместите Excel файлы в папку Source и запустите программу снова.');
            // Приостанавливаем выполнение, чтобы пользователь успел прочитать сообщение
            console.log('Нажмите Enter для завершения...');
            // В исполняемом файле просто ждем несколько секунд
            setTimeout(() => process.exit(0), 10000);
            return [];
        } catch (error) {
            console.error(`Ошибка при создании папки Source: ${error.message}`);
            console.log('Нажмите Enter для завершения...');
            setTimeout(() => process.exit(1), 10000);
            return [];
        }
    }

    let allFiles = [];

    sourceDirs.forEach(dir => {
        console.log(`Ищем Excel файлы в: ${dir}`);
        try {
            const files = fs.readdirSync(dir);
            const excelFiles = files.filter(file =>
                path.extname(file).toLowerCase() === '.xlsx' &&
                !file.startsWith('~$') // Исключаем временные файлы Excel
            ).map(file => ({
                name: file,
                path: path.join(dir, file)
            }));

            allFiles = [...allFiles, ...excelFiles];
        } catch (error) {
            console.error(`Ошибка при чтении директории ${dir}: ${error.message}`);
        }
    });

    console.log(`Найдено ${allFiles.length} Excel файлов.`);
    return allFiles;
}

// Function to trim country name to 2 characters (for columns 7 and 8)
function trimCountryName(value, columnIndex) {
    // Проверяем, что это 7-я или 8-я колонка (индексы 6 и 7, т.к. нумерация с 0)
    if (columnIndex === 6 || columnIndex === 7) {
        if (typeof value === 'string' && value.length > 2) {
            // Возвращаем только первые два символа
            return value.substring(0, 2);
        }
    }
    return value;
}

// Function to reorder columns based on first row values
function reorderColumns(data) {
    if (!data || data.length < 2) return data; // Нужно как минимум заголовок и одна строка данных

    const firstRow = data[0];

    // Проверяем первую строку, ищем числовые значения для порядка колонок
    const columnOrder = firstRow.map((value, index) => {
        // Преобразуем значение в число, если это возможно
        const orderValue = parseInt(value);
        // Возвращаем объект с индексом колонки и порядковым значением
        return {
            index: index,
            order: isNaN(orderValue) ? -1 : orderValue,
            header: value
        };
    });

    // Фильтруем колонки с порядковым номером 0 (которые нужно пропустить)
    const validColumns = columnOrder.filter(col => col.order !== 0);

    // Если нет действительных колонок для перестановки, возвращаем исходные данные
    if (validColumns.length === 0) return data;

    // Получаем только колонки с числовыми значениями (положительными)
    const numberedColumns = validColumns.filter(col => col.order > 0);

    // Находим максимальный номер колонки
    const maxColumnOrder = numberedColumns.reduce((max, col) => Math.max(max, col.order), 0);

    // Создаем массив для хранения результата с непрерывной нумерацией
    // Индекс массива = порядковый номер - 1 (так как порядок начинается с 1)
    const orderedColumnsMap = new Array(maxColumnOrder).fill(null).map(() => ({
        indices: [], // Массив индексов колонок с таким номером (для дублирующихся)
        header: '' // Заголовок колонки
    }));

    // Заполняем orderedColumnsMap
    numberedColumns.forEach(col => {
        const index = col.order - 1; // Индекс в массиве (порядковый номер - 1)
        orderedColumnsMap[index].indices.push(col.index);

        // Если заголовок уже есть и новый заголовок отличается, добавляем его через разделитель
        if (orderedColumnsMap[index].header && orderedColumnsMap[index].header !== col.header) {
            orderedColumnsMap[index].header += ' / ' + col.header;
        } else {
            orderedColumnsMap[index].header = col.header;
        }
    });

    // Добавляем колонки без порядкового номера в конец
    const columnsWithoutOrder = validColumns.filter(col => col.order === -1);

    // Создаем новый массив данных с переставленными колонками и пустыми столбцами
    const newData = [];

    // Создаем новые заголовки с правильной нумерацией
    const newHeaders = [];

    // Добавляем заголовки для колонок с нумерацией
    orderedColumnsMap.forEach(column => {
        newHeaders.push(column.header || ''); // Если header пустой, добавляем пустую строку
    });

    // Добавляем заголовки для колонок без нумерации
    columnsWithoutOrder.forEach(col => {
        newHeaders.push(firstRow[col.index] || '');
    });

    // Добавляем заголовки в результат
    newData.push(newHeaders);

    // Начинаем добавлять данные с третьей строки (индекс 2), пропуская первые две строки
    for (let i = 3; i < data.length; i++) {
        const row = data[i];
        const newRow = [];

        // Добавляем ячейки для колонок с нумерацией
        orderedColumnsMap.forEach((column, colIndex) => {
            if (column.indices.length === 0) {
                // Если у этого номера нет соответствующих колонок, добавляем пустую ячейку
                newRow.push('');
            } else if (column.indices.length === 1) {
                // Если только одна колонка с таким номером
                let value = row[column.indices[0]] || '';
                // Обрезаем названия стран в 7-й и 8-й колонках
                value = trimCountryName(value, colIndex);
                newRow.push(value);
            } else {
                // Если несколько колонок с таким номером, конкатенируем значения
                const values = column.indices
                    .map(idx => {
                        let value = row[idx] || '';
                        // Обрезаем названия стран в 7-й и 8-й колонках
                        value = trimCountryName(value, colIndex);
                        return value;
                    })
                    .filter(val => val !== ''); // Фильтруем пустые значения

                newRow.push(values.join(' / '));
            }
        });

        // Добавляем ячейки для колонок без нумерации
        columnsWithoutOrder.forEach(col => {
            newRow.push(row[col.index] || '');
        });

        newData.push(newRow);
    }

    return newData;
}

// Функция для создания сводной таблицы из данных
function createPivotTable(data) {
    if (!data || data.length < 2) return []; // Нужно как минимум заголовок и одна строка данных

    const headers = data[0];
    const rows = data.slice(1);

    // Индексы полей для группировки (3, 17, 12) - нужно уменьшить на 1 для индексации в массиве
    const groupByIndices = [2, 16, 11];
    // Индексы полей для конкатенации через запятую (8, 11, 15, 16, 6) - уменьшаем на 1
    const concatIndices = [7, 10, 14, 15, 5];

    // Создаем объект для хранения сгруппированных данных
    const groupedData = {};

    // Группируем данные
    rows.forEach(row => {
        // Создаем ключ группировки на основе полей 3, 17, 12
        const groupKey = groupByIndices.map(idx => row[idx] || '').join('|');

        if (!groupedData[groupKey]) {
            // Инициализируем группу
            groupedData[groupKey] = {
                count: 0,
                rows: [],
                concatValues: groupByIndices.map(idx => ({
                    colIndex: idx,
                    value: row[idx] || ''
                }))
            };
        }

        // Увеличиваем счетчик записей в группе
        groupedData[groupKey].count++;
        // Добавляем строку в группу
        groupedData[groupKey].rows.push(row);
    });

    // Определяем все индексы столбцов
    const allColumnIndices = Array.from({ length: headers.length }, (_, i) => i);

    // Находим индексы для "остальных" полей (не группировка и не конкатенация)
    const regularIndices = allColumnIndices.filter(idx =>
        !groupByIndices.includes(idx) && !concatIndices.includes(idx)
    );

    // Создаем массив с новым порядком столбцов:
    // 1. Поля группировки (3, 17, 12)
    // 2. Количество записей (будет добавлено позже)
    // 3. Поля конкатенации (8, 11, 15, 16, 6)
    // 4. Все остальные поля
    const columnOrder = [
        ...groupByIndices,
        -1, // Место для столбца количества (временный маркер)
        ...concatIndices,
        ...regularIndices
    ];

    // Создаем новые заголовки для сводной таблицы с маркировкой, уже в нужном порядке
    const pivotHeaders = columnOrder.map(idx => {
        if (idx === -1) {
            return 'кол-во записей';
        } else if (idx === 2) { // 3-й столбец (индекс 2)
            return 'компания';
        } else if (idx === 16) { // 17-й столбец (индекс 16)
            return 'Инкотермс';
        } else if (idx === 11) { // 12-й столбец (индекс 11)
            return 'таможня назначения';
        } else if (idx === 7) { // 8-й столбец (индекс 7)
            return 'страны отправления';
        } else if (idx === 10) { // 11-й столбец (индекс 10)
            return 'страны назначения';
        } else if (idx === 14) { // 15-й столбец (индекс 14)
            return 'коды ТНВЭД';
        } else if (idx === 15) { // 16-й столбец (индекс 15)
            return 'товарные знаки';
        } else if (idx === 5) { // 6-й столбец (индекс 5)
            return 'контактная информация';
        } else if (groupByIndices.includes(idx)) {
            return `[ГРУППА] ${headers[idx]}`;
        } else if (concatIndices.includes(idx)) {
            return `[КОНКАТ] ${headers[idx]}`;
        } else {
            return `[ПЕРВЫЙ] ${headers[idx]}`;
        }
    });

    // Создаем структуру для сводной таблицы
    const pivotData = [pivotHeaders];

    // Формируем строки сводной таблицы
    Object.keys(groupedData).forEach(groupKey => {
        const group = groupedData[groupKey];

        // Создаем новую строку для сводной таблицы в соответствии с новым порядком
        const pivotRow = [];

        // Заполняем строку в соответствии с новым порядком столбцов
        columnOrder.forEach(idx => {
            // Для столбца количества
            if (idx === -1) {
                pivotRow.push(group.count.toString());
                return;
            }

            // Для полей группировки
            if (groupByIndices.includes(idx)) {
                const groupIndex = groupByIndices.indexOf(idx);
                pivotRow.push(group.concatValues[groupIndex].value);
                return;
            }

            // Для полей конкатенации
            if (concatIndices.includes(idx)) {
                // Получаем уникальные значения
                const uniqueValues = new Set();
                group.rows.forEach(row => {
                    if (row[idx] && row[idx] !== '') {
                        uniqueValues.add(row[idx]);
                    }
                });
                // Конкатенируем через запятую
                pivotRow.push(Array.from(uniqueValues).join(', '));
                return;
            }

            // Для остальных полей - берем первое непустое значение
            let value = '';
            for (const row of group.rows) {
                if (row[idx] && row[idx] !== '') {
                    value = row[idx];
                    break;
                }
            }
            pivotRow.push(value);
        });

        // Добавляем строку в сводную таблицу
        pivotData.push(pivotRow);
    });

    return pivotData;
}

// Function to merge Excel files
function mergeExcelFiles() {
    const excelFiles = readExcelFiles();
    if (excelFiles.length === 0) {
        console.log("Не найдено Excel файлов для объединения.");
        console.log('Нажмите Enter для завершения...');
        setTimeout(() => process.exit(0), 10000);
        return;
    }

    console.log(`Found ${excelFiles.length} Excel files to merge.`);

    // Create a new workbook
    const mergedWorkbook = XLSX.utils.book_new();

    // Массив для хранения всех данных из всех файлов
    let allData = [];
    // Флаг для отслеживания, были ли добавлены заголовки
    let headersAdded = false;

    // Process each Excel file
    excelFiles.forEach((file, index) => {
        console.log(`Processing file ${index + 1}/${excelFiles.length}: ${file.name}`);
        const filePath = file.path;

        try {
            // Read the workbook
            const workbook = XLSX.readFile(filePath);

            // Process each sheet in the workbook
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];

                // Получаем диапазон ячеек
                const range = XLSX.utils.decode_range(worksheet['!ref']);

                // Если есть данные
                if (range.e.r >= range.s.r) {
                    // Получаем данные в виде массива массивов
                    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (data.length > 0) {
                        // Переставляем колонки на основе порядковых номеров в первой строке
                        const reorderedData = reorderColumns(data);

                        // Если это первый файл и лист, добавляем заголовки
                        if (!headersAdded) {
                            // Добавляем заголовки (первую строку)
                            allData.push(reorderedData[0]);
                            headersAdded = true;
                        }

                        // Добавляем данные (все строки кроме заголовка)
                        if (reorderedData.length > 1) {
                            allData = [...allData, ...reorderedData.slice(1)];
                        }
                    }
                }
            });
        } catch (error) {
            console.error(`Error processing file ${file.name}:`, error);
        }
    });

    // Если есть данные, создаем листы
    if (allData.length > 0) {

        const originalHeaders = allData[0];

        //
        // Маппинг номеров полей на их названия (согласно картинке)
        const fieldMappings = {
            1: 'отправитель',
            2: 'адрес отправителя',
            3: 'контрактодержатель',
            4: 'адрес контарктодерж',
            5: 'ИНН',
            6: 'контактная информация',
            7: 'торгующая страна',
            8: 'страна отправления',
            9: 'страна происхождения',
            10: 'производитель',
            11: 'страна назначения',
            12: 'таможня назначения',
            13: 'тип транспорта',
            14: 'описание товара',
            15: 'код ТНВЭД',
            16: 'товарный знак',
            17: 'Инкотермс',
            18: 'пункт назначения',
            19: 'вес брутто',
            20: 'кол-во',
            21: 'ед изм',
            22: 'факт стоим',
            23: 'валюта'
        };

        // Создаем новые заголовки на основе номеров в первой строке
        const newHeaders = originalHeaders.map((header, index) => {
            // Пытаемся извлечь числовое значение из заголовка
            const orderValue = parseInt(header);

            // Если это число и оно присутствует в нашем маппинге
            if (!isNaN(orderValue) && fieldMappings[orderValue]) {
                return `${fieldMappings[orderValue]}`;
            }
            // Если это число, но нет в маппинге
            else if (!isNaN(orderValue)) {
                return `${orderValue}`;
            }
            // Если это не число, оставляем как есть
            else {
                return header;
            }
        });

        // Заменяем первую строку новыми заголовками
        allData[0] = newHeaders;

        // Лист 1: Все объединенные данные
        const mergedWorksheet = XLSX.utils.aoa_to_sheet(allData);
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedWorksheet, "Merged Data");

        // Лист 2: Сводная таблица
        const pivotData = createPivotTable(allData);
        if (pivotData.length > 1) { // Если есть данные для сводной таблицы
            const pivotWorksheet = XLSX.utils.aoa_to_sheet(pivotData);

            // Устанавливаем ширину столбцов
            if (!pivotWorksheet['!cols']) pivotWorksheet['!cols'] = [];
            for (let i = 0; i < pivotData[0].length; i++) {
                pivotWorksheet['!cols'][i] = { width: 20 }; // Увеличиваем ширину для маркированных заголовков
            }

            XLSX.utils.book_append_sheet(mergedWorkbook, pivotWorksheet, "Pivot Table");
        }

        // Write the merged workbook to a file
        try {
            XLSX.writeFile(mergedWorkbook, outputFile);
            console.log(`Successfully merged Excel files. Output saved to: ${outputFile}`);
            // Добавляем паузу в конце программы, чтобы пользователь мог прочитать сообщение
            console.log('Нажмите Enter для завершения...');
            setTimeout(() => process.exit(0), 10000);
        } catch (error) {
            console.error('Error writing merged file:', error);
            console.log('Возможно, у программы нет прав на запись в эту директорию.');
            console.log(`Попробуйте запустить программу от имени администратора или скопировать её в другую папку.`);
            console.log('Нажмите Enter для завершения...');
            setTimeout(() => process.exit(1), 10000);
        }
    } else {
        console.log("No data found to merge.");
    }
}

// Execute the merge function
mergeExcelFiles(); 