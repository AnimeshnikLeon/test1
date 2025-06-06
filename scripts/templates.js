(function(window, undefined){  
  
	function updateScroll()  
	{  
		Ps.update();  
	}  
  
	var _templates = [  
		[ "bold_report", 122, 158 ],  
		[ "calendar_2025", 158, 122 ],  
		[ "izav_import", 158, 122 ],  
	];  
  
	var _templates_code = [];  
  
	function fill_templates()  
	{  
		var _width = 0;  
		for (var i = 0; i < _templates.length; i++)  
		{  
			if (_templates[i][1] > _width)  
				_width = _templates[i][1];  
		}  
  
		_width += 20;  
  
		var _space = 20;  
		var _naturalWidth = window.innerWidth;  
  
		var _count = ((_naturalWidth - _space) / (_width + _space)) >> 0;  
		if (_count < 1)  
			_count = 1;  
  
		var _countRows = ((_templates.length + (_count - 1)) / _count) >> 0;  
  
		var _html = "";  
		var _index = 0;  
  
		var _margin = (_naturalWidth - _count * (_width + _space)) >> 1;  
		document.getElementById("main").style.marginLeft = _margin + "px";  
  
		for (var _row = 0; _row < _countRows && _index < _templates.length; _row++)  
		{  
			_html += "<tr style='margin-left: " + _margin + "'>";  
  
			for (var j = 0; j < _count; j++)  
			{  
				var _cur = _templates[_index];  
  
				_html += "<td width='" + _width + "' height='" +_width + "' style='margin:" + (_space >> 1) + "'>";  
  
				var _w = _cur[1];  
				var _h = _cur[2];  
  
				_html += ("<img id='template" + _index + "' src=\"./templates/" + _cur[0] + "/icon.png\" />");  
				_html += ("<div class=\"noselect celllabel\">" + _cur[0] + "</div>");  
  
				_html += "</td>";  
  
				_index++;  
  
				if (_index >= _templates.length)  
					break;  
			}  
  
			_html += "</tr>";  
		}  
  
		document.getElementById("main").innerHTML = _html;  
  
		for (_index = 0; _index < _templates.length; _index++)  
		{  
			document.getElementById("template" + _index).onclick = new Function("return window.template_run(" + _index + ");");  
		}  
  
		updateScroll();  
	}  
  
	window.onresize = function()  
	{  
		fill_templates();  
	};  
  
	window.Asc.plugin.init = function()  
	{  
		var container = document.getElementById('scrollable-container-id');  
		  
		Ps = new PerfectScrollbar('#' + container.id, {});  
  
		fill_templates();  
	};  
	  
	window.Asc.plugin.button = function(id)  
	{  
		this.executeCommand("close", "");  
	};  
  
	// Функция показа интерфейса загрузки файла  
	window.show_file_upload = function() {  
		document.getElementById('scrollable-container-id').style.display = 'none';  
		document.getElementById('file-upload-container').style.display = 'block';  
		  
		// Обработчик загрузки файла  
		document.getElementById('process-file').onclick = function() {  
			const fileInput = document.getElementById('xlsx-file');  
			const sheetName = document.getElementById('sheet-name').value;  
			const statusDiv = document.getElementById('status-message');  
			  
			if (!fileInput.files[0]) {  
				statusDiv.innerHTML = '<span style="color: red;">Пожалуйста, выберите файл</span>';  
				return;  
			}  
			  
			if (!sheetName.trim()) {  
				statusDiv.innerHTML = '<span style="color: red;">Пожалуйста, укажите имя листа</span>';  
				return;  
			}  
			  
			statusDiv.innerHTML = '<span style="color: blue;">Обработка файла...</span>';  
			processXLSXFile(fileInput.files[0], sheetName.trim());  
		};  
		  
		// Кнопка возврата к шаблонам  
		document.getElementById('back-to-templates').onclick = function() {  
			document.getElementById('file-upload-container').style.display = 'none';  
			document.getElementById('scrollable-container-id').style.display = 'block';  
			document.getElementById('xlsx-file').value = '';  
			document.getElementById('sheet-name').value = '';  
			document.getElementById('status-message').innerHTML = '';  
		};  
	};  
  
	// Функция обработки XLSX файла с правильной логикой парсинга ИЗАВ  
	function processXLSXFile(file, sheetName) {  
		const statusDiv = document.getElementById('status-message');  
		
		const reader = new FileReader();  
		reader.onload = function(e) {  
			try {  
				const data = new Uint8Array(e.target.result);  
				const workbook = XLSX.read(data, {type: 'array'});  
				
				// Проверяем существование листа  
				if (!workbook.SheetNames.includes(sheetName)) {  
					statusDiv.innerHTML = '<span style="color: red;">Лист "' + sheetName + '" не найден. Доступные листы: ' + workbook.SheetNames.join(', ') + '</span>';  
					return;  
				}  
				
				const worksheet = workbook.Sheets[sheetName];  
				const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false, defval: ""});  
				
				statusDiv.innerHTML = '<span style="color: blue;">Парсинг данных...</span>';  
				
				// Парсим данные согласно требованиям ИЗАВ  
				const parsedData = parseIZAVData(jsonData);  
				
				statusDiv.innerHTML = '<span style="color: green;">Обработано строк: ' + parsedData.length + '</span>';  
				
				// Создаем DOCX документ  
				createIZAVDocument(parsedData);  
				
			} catch (error) {  
				statusDiv.innerHTML = '<span style="color: red;">Ошибка обработки файла: ' + error.message + '</span>';  
			}  
		};  
		
		reader.readAsArrayBuffer(file);  
	}  
	
	// Функция парсинга данных ИЗАВ согласно требованиям  
	function parseIZAVData(rawData) {  
		const result = [];  
		
		for (let i = 0; i < rawData.length; i++) {  
			const row = rawData[i];  
			
			// Проверяем конец данных  
			if (row[1] === "EOD") {  
				break;  
			}  
			
			// Проверяем критерий включения строки (номера в столбцах A и E)  
			if (shouldIncludeRow(row)) {  
				// Обрабатываем строку данных  
				const processedRows = processDataRow(row);  
				result.push(...processedRows);  
			} else if (isHeaderRow(row)) {  
				// Добавляем строки заголовков как есть  
				result.push(row);  
			}  
		}  
		
		return result;  
	}  
	
	// Упрощенная проверка строки заголовка  
	function isHeaderRow(row) {  
		// Строка заголовка содержит текст в первой ячейке, но не является строкой данных  
		if (!row[0] || row[0].trim() === "") return false;  
		
		// Если в столбце A есть число 1-2 знака И в столбце E есть 4-значное число - это данные  
		const columnA = String(row[0] || "").trim();  
		const columnE = String(row[4] || "").trim(); // Столбец E (индекс 4)  
		
		const hasValidA = /^\d{1,2}$/.test(columnA);  
		const hasValidE = /^\d{4}$/.test(columnE);  
		
		// Если это не строка данных, то это заголовок  
		return !(hasValidA && hasValidE);  
	}  
	
	// Исправленная проверка критерия включения строки  
	function shouldIncludeRow(row) {  
		// Критерий: наличие номеров в столбце A (1-2 знака) и в столбце E (4 знака)  
		const columnA = String(row[0] || "").trim();  
		const columnE = String(row[4] || "").trim(); // Столбец E (индекс 4)  
		
		const hasValidA = /^\d{1,2}$/.test(columnA);  
		const hasValidE = /^\d{4}$/.test(columnE);  
		
		return hasValidA && hasValidE;  
	}  
	
	// Исправленная обработка строки данных  
	function processDataRow(row) {  
		const processedRows = [];  
		
		// Извлекаем координаты из столбцов H, I, J, K, L, M (индексы 7, 8, 9, 10, 11, 12)  
		const coordinates = {  
			H: parseCoordinates(row[7]),  // WGS-84 широта  
			I: parseCoordinates(row[8]),  // WGS-84 долгота    
			J: parseCoordinates(row[9]),  // ГСК-2011 X  
			K: parseCoordinates(row[10]), // ГСК-2011 Y  
			L: parseCoordinates(row[11]), // МСК-26 X  
			M: parseCoordinates(row[12])  // МСК-26 Y  
		};  
		
		// Определяем максимальное количество координат  
		const maxCoords = Math.max(  
			coordinates.H.length,  
			coordinates.I.length,  
			coordinates.J.length,  
			coordinates.K.length,  
			coordinates.L.length,  
			coordinates.M.length,  
			1  
		);  
		
		// Создаем строки для каждого набора координат  
		for (let coordIndex = 0; coordIndex < maxCoords; coordIndex++) {  
			const newRow = [...row];  
			
			// Заполняем координаты с правильным форматированием  
			newRow[7] = formatCoordinate(coordinates.H[coordIndex], "WGS84_LAT");  
			newRow[8] = formatCoordinate(coordinates.I[coordIndex], "WGS84_LON");  
			newRow[9] = formatCoordinate(coordinates.J[coordIndex], "GSK2011");  
			newRow[10] = formatCoordinate(coordinates.K[coordIndex], "GSK2011");  
			newRow[11] = formatCoordinate(coordinates.L[coordIndex], "MSK26");  
			newRow[12] = formatCoordinate(coordinates.M[coordIndex], "MSK26");  
			
			// Проверяем ошибки  
			newRow._errors = checkDataErrors(newRow);  
			
			processedRows.push(newRow);  
		}  
		
		return processedRows;  
	}  
	
	// Исправленное форматирование координат  
	function formatCoordinate(value, type) {  
		if (!value || value === "" || value === undefined) return "";  
		
		// Убираем лишние пробелы и заменяем запятые на точки  
		const cleanValue = value.toString().trim().replace(',', '.');  
		const num = parseFloat(cleanValue);  
		
		if (isNaN(num)) return cleanValue;  
		
		switch (type) {  
			case "WGS84_LAT":  
			case "WGS84_LON":  
				// 2 знака до запятой, 9 после (например: 44.915841889)  
				return num.toFixed(9);  
			case "GSK2011":  
				// 6-8 знаков до запятой, 3 после (например: 8482277.107)  
				return num.toFixed(3);  
			case "MSK26":  
				// 6-7 знаков до запятой, 3 после (например: 490433.357)  
				return num.toFixed(3);  
			default:  
				return cleanValue;  
		}  
	}
	
	// Парсит координаты из ячейки (может содержать 1, 2 или 4 числа через пробелы)  
	function parseCoordinates(cellValue) {  
		if (!cellValue || cellValue.toString().trim() === "") {  
			return [""];  
		}  
		
		const str = cellValue.toString().trim();  
		// Разделяем по группам пробелов  
		const coords = str.split(/\s+/).filter(val => val !== "");  
		
		return coords.length > 0 ? coords : [""];  
	}

	// Проверяет ошибки в данных  
	function checkDataErrors(row) {  
		const errors = [];  
		
		// Проверяем обязательные координаты H, I  
		if (!row[7] || row[7].trim() === "") {  
			errors.push("H");  
		}  
		if (!row[8] || row[8].trim() === "") {  
			errors.push("I");  
		}  
		
		// Проверяем обязательные координаты L, M  
		if (!row[11] || row[11].trim() === "") {  
			errors.push("L");  
		}  
		if (!row[12] || row[12].trim() === "") {  
			errors.push("M");  
		}  
		
		// J и K могут быть пустыми (это нормально)  
		// O может быть пустым (это нормально)  
		
		return errors;  
	}
  
	// Функция создания DOCX документа с правильным форматированием
	function createIZAVDocument(data) {  
		const script = `  
			var oDocument = Api.GetDocument();  
			
			// Создаем таблицу с 10 столбцами (как в HTML)  
			var oTable = Api.CreateTable(10, ${data.length + 3});  
			
			// Настройки таблицы  
			oTable.SetTableLayout("autofit");  
			oTable.SetTableLook(true, true, false, false, true, false);  
			
			// Заголовки первой строки (с объединением ячеек)  
			var headerRow1 = [  
				"№ п/п",           // rowspan=2  
				"Источник выделения", // rowspan=2    
				"Номер ИЗАВ",      // rowspan=2  
				"WGS-84",          // colspan=2  
				"",                // объединена с предыдущей  
				"ГСК-2011",        // colspan=2  
				"",                // объединена с предыдущей  
				"МСК-26 от СК-95 (зона 2)", // colspan=2  
				"",                // объединена с предыдущей  
				"Метод инвентаризации выбросов" // rowspan=2  
			];  
			
			// Заголовки второй строки  
			var headerRow2 = [  
				"", "", "",        // пустые (rowspan из первой строки)  
				"Северная широта",  
				"Восточная долгота",   
				"Северная широта",  
				"Восточная долгота",  
				"координата X",  
				"координата Y",  
				""                 // пустая (rowspan из первой строки)  
			];  
			
			// Номера столбцов (третья строка заголовков)  
			var columnNumbers = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];  
			
			// Заполняем первую строку заголовков  
			for (var i = 0; i < headerRow1.length; i++) {  
				if (headerRow1[i] !== "") {  
					var oCell = oTable.GetRow(0).GetCell(i);  
					var oParagraph = oCell.GetContent().GetElement(0);  
					oParagraph.AddText(headerRow1[i]);  
					oParagraph.SetJc("center");  
					
					// Границы ячейки  
					oCell.SetCellBorderTop("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderBottom("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderLeft("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderRight("single", 1, 0, 0, 0, 0);  
					
					var oTextPr = oParagraph.GetTextPr();  
					oTextPr.SetFontFamily("Calibri");  
					oTextPr.SetFontSize(18); // 9pt = 18 в единицах API  
					oTextPr.SetBold(true);  
				}  
			}  
			
			// Заполняем вторую строку заголовков  
			for (var i = 0; i < headerRow2.length; i++) {  
				if (headerRow2[i] !== "") {  
					var oCell = oTable.GetRow(1).GetCell(i);  
					var oParagraph = oCell.GetContent().GetElement(0);  
					oParagraph.AddText(headerRow2[i]);  
					oParagraph.SetJc("center");  
					
					// Границы ячейки  
					oCell.SetCellBorderTop("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderBottom("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderLeft("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderRight("single", 1, 0, 0, 0, 0);  
					
					var oTextPr = oParagraph.GetTextPr();  
					oTextPr.SetFontFamily("Calibri");  
					oTextPr.SetFontSize(18);  
					oTextPr.SetBold(true);  
				}  
			}  
			
			// Заполняем третью строку (номера столбцов)  
			for (var i = 0; i < columnNumbers.length; i++) {  
				var oCell = oTable.GetRow(2).GetCell(i);  
				var oParagraph = oCell.GetContent().GetElement(0);  
				oParagraph.AddText(columnNumbers[i]);  
				oParagraph.SetJc("center");  
				
				// Границы ячейки  
				oCell.SetCellBorderTop("single", 1, 0, 0, 0, 0);  
				oCell.SetCellBorderBottom("single", 1, 0, 0, 0, 0);  
				oCell.SetCellBorderLeft("single", 1, 0, 0, 0, 0);  
				oCell.SetCellBorderRight("single", 1, 0, 0, 0, 0);  
				
				var oTextPr = oParagraph.GetTextPr();  
				oTextPr.SetFontFamily("Calibri");  
				oTextPr.SetFontSize(18);  
				oTextPr.SetBold(true);  
			}  
			
			// Заполняем данные  
			var dataArray = ${JSON.stringify(data)};  
			for (var row = 0; row < dataArray.length; row++) {  
				var rowData = dataArray[row];  
				
				// Проверяем, является ли строка заголовком (объединенные ячейки)  
				var isHeaderRow = isFullRowHeader(rowData);  
				
				if (isHeaderRow) {  
					// Объединенная строка заголовка на всю ширину  
					var oCell = oTable.GetRow(row + 3).GetCell(0);  
					var oParagraph = oCell.GetContent().GetElement(0);  
					oParagraph.AddText(String(rowData[0] || ""));  
					oParagraph.SetJc("center");  
					
					// Границы и форматирование  
					oCell.SetCellBorderTop("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderBottom("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderLeft("single", 1, 0, 0, 0, 0);  
					oCell.SetCellBorderRight("single", 1, 0, 0, 0, 0);  
					
					var oTextPr = oParagraph.GetTextPr();  
					oTextPr.SetFontFamily("Calibri");  
					oTextPr.SetFontSize(18);  
					oTextPr.SetBold(true);  
				} else {  
					// Обычная строка данных  
					for (var col = 0; col < Math.min(rowData.length, 10); col++) {  
						var oCell = oTable.GetRow(row + 3).GetCell(col);  
						var oParagraph = oCell.GetContent().GetElement(0);  
						var cellValue = formatCellValue(rowData[col] || "", col);  
						oParagraph.AddText(String(cellValue));  
						
						// Выравнивание согласно HTML: столбец 1 (Источник выделения) - влево, остальные - по центру  
						if (col === 1) {  
							oParagraph.SetJc("left");  
						} else {  
							oParagraph.SetJc("center");  
						}  
						
						// Границы ячейки  
						oCell.SetCellBorderTop("single", 1, 0, 0, 0, 0);  
						oCell.SetCellBorderBottom("single", 1, 0, 0, 0, 0);  
						oCell.SetCellBorderLeft("single", 1, 0, 0, 0, 0);  
						oCell.SetCellBorderRight("single", 1, 0, 0, 0, 0);  
						
						var oTextPr = oParagraph.GetTextPr();  
						oTextPr.SetFontFamily("Calibri");  
						oTextPr.SetFontSize(18);  
						
						// Оранжевая заливка для ошибок  
						if (rowData._errors && shouldHighlightError(col, rowData._errors)) {  
							oCell.SetShd("clear", 255, 192, 0, false); // #FFC000  
						}  
					}  
				}  
			}  
			
			// Функция форматирования значений согласно HTML спецификации  
			function formatCellValue(value, columnIndex) {  
				if (!value || value === "") return "";  
				
				var strValue = String(value);  
				
				// Столбец 2 (Номер ИЗАВ) - формат 'хххх'  
				if (columnIndex === 2) {  
					var num = parseInt(strValue);  
					if (!isNaN(num)) {  
						return String(num).padStart(4, '0');  
					}  
				}  
				
				// Столбцы 3,4 (WGS-84) - 2 знака до запятой, 9 после  
				if (columnIndex === 3 || columnIndex === 4) {  
					var num = parseFloat(strValue.replace(',', '.'));  
					if (!isNaN(num)) {  
						return num.toFixed(9);  
					}  
				}  
				
				// Столбцы 5,6 (ГСК-2011) - 6-8 знаков до запятой, 3 после  
				if (columnIndex === 5 || columnIndex === 6) {  
					var num = parseFloat(strValue.replace(',', '.'));  
					if (!isNaN(num)) {  
						return num.toFixed(3);  
					}  
				}  
				
				// Столбцы 7,8 (МСК-26) - 6-7 знаков до запятой, 3 после  
				if (columnIndex === 7 || columnIndex === 8) {  
					var num = parseFloat(strValue.replace(',', '.'));  
					if (!isNaN(num)) {  
						return num.toFixed(3);  
					}  
				}  
				
				return value;  
			}  
			
			// Проверка на строку заголовка (объединенные ячейки)  
			function isFullRowHeader(row) {  
				if (!row || !row[0]) return false;  
				// Если есть текст в первой ячейке, но нет номеров в A и F - это заголовок  
				var columnA = String(row[0] || "").trim();  
				var columnF = String(row[5] || "").trim();  
				var hasValidA = /^\\d{1,2}$/.test(columnA);  
				var hasValidF = /^\\d{4}$/.test(columnF);  
				return !hasValidA && !hasValidF && columnA !== "";  
			}  
			
			// Функция для определения оранжевой заливки  
			function shouldHighlightError(colIndex, errors) {  
				var columnMap = {3: "H", 4: "I", 7: "L", 8: "M"};  
				var columnLetter = columnMap[colIndex];  
				return columnLetter && errors.includes(columnLetter);  
			}  
			
			oDocument.Push(oTable);  
		`;  
		
		window.Asc.plugin.info.recalculate = true;  
		window.Asc.plugin.executeCommand("command", script);  
	}
	
	window.template_run = function(_index)  
	{  
		// Специальная обработка для шаблона импорта ИЗАВ  
		if (_templates[_index][0] === "izav_import") {  
			window.show_file_upload();  
			return;  
		}  
		  
		// Стандартная обработка для остальных шаблонов  
		if (_templates_code[_index])  
		{  
			window.Asc.plugin.info.recalculate = true;  
			window.Asc.plugin.executeCommand("command", _templates_code[_index]);  
			return;  
		}  
  
		window.Asc.plugin.callModule("./templates/" + _templates[_index][0] + "/script.txt", function(content){  
			_templates_code[_index] = content;  
			window.Asc.plugin.info.recalculate = true;  
			window.Asc.plugin.executeCommand("command", content);  
		});  
	};  
  
	window.Asc.plugin.onExternalMouseUp = function()  
    {  
        var evt = document.createEvent("MouseEvents");  
        evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,  
            false, false, false, false, 0, null);  
  
        document.dispatchEvent(evt);  
    };  
  
})(window, undefined);