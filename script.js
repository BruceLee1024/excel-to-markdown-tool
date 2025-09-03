document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excelFile');
    const convertBtn = document.getElementById('convertBtn');
    const markdownOutputTextarea = document.getElementById('markdownOutput');
    const copyBtn = document.getElementById('copyBtn');
    const exportBtn = document.getElementById('exportBtn');
    const fileList = document.getElementById('fileList');
    const columnSelector = document.getElementById('columnSelector');
    const columnList = document.getElementById('columnList');
    const selectAllColumns = document.getElementById('selectAllColumns');
    const applyColumnFilter = document.getElementById('applyColumnFilter');
    const cancelColumnFilter = document.getElementById('cancelColumnFilter');
    const selectColumnsBtn = document.getElementById('selectColumnsBtn');
    const mergeExcelBtn = document.getElementById('mergeExcelBtn');
    const mergeToCsvBtn = document.getElementById('mergeToCsvBtn');
    
    let selectedFiles = [];
    let currentWorkbookData = null;
    let selectedColumns = new Set();
    let mergedWorkbook = null;
    
    // 文件选择处理
    excelFileInput.addEventListener('change', (event) => {
        const files = Array.from(event.target.files);
        selectedFiles = [...selectedFiles, ...files];
        updateFileList();
        updateButtonStates();
    });
    
    // 更新文件列表显示
    function updateFileList() {
        fileList.innerHTML = '';
        selectedFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <div class="file-info">
                    <svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14,2 14,8 20,8"></polyline>
                        <line x1="16" y1="13" x2="8" y2="13"></line>
                        <line x1="16" y1="17" x2="8" y2="17"></line>
                        <polyline points="10,9 9,9 8,9"></polyline>
                    </svg>
                    <span class="file-name">${file.name}</span>
                    <span class="file-size">(${formatFileSize(file.size)})</span>
                </div>
                <button class="remove-file" onclick="removeFile(${index})">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
            `;
            fileList.appendChild(fileItem);
        });
    }
    
    // 移除文件
    window.removeFile = function(index) {
        selectedFiles.splice(index, 1);
        updateFileList();
        updateButtonStates();
    };
    
    // 格式化文件大小
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // 更新按钮状态
    function updateButtonStates() {
        const hasFiles = selectedFiles.length > 0;
        const hasOutput = markdownOutputTextarea.value.trim() !== '';
        const hasSingleFile = selectedFiles.length === 1;
        const hasMultipleFiles = selectedFiles.length > 1;
        
        convertBtn.disabled = !hasFiles;
        convertBtn.title = hasFiles ? '开始转换Excel文件' : '需要上传至少1个Excel文件';
        
        selectColumnsBtn.disabled = !hasSingleFile;
        selectColumnsBtn.title = hasSingleFile ? '选择要保留的列' : '需要上传1个Excel文件';
        
        mergeExcelBtn.disabled = !hasMultipleFiles;
        mergeExcelBtn.title = hasMultipleFiles ? '合并多个Excel文件后转换' : '需要上传至少2个Excel文件';
        
        mergeToCsvBtn.disabled = !hasMultipleFiles;
        mergeToCsvBtn.title = hasMultipleFiles ? '合并多个Excel文件后转换为CSV格式' : '需要上传至少2个Excel文件';
        
        copyBtn.disabled = !hasOutput;
        exportBtn.disabled = !hasOutput;
    }
    
    // 选择列后转换按钮事件
    selectColumnsBtn.addEventListener('click', async () => {
        if (selectedFiles.length !== 1) {
            alert('列选择功能仅支持单个文件！');
            return;
        }
        
        try {
            const workbook = await loadWorkbook(selectedFiles[0]);
            currentWorkbookData = workbook;
            showColumnSelector(workbook);
        } catch (error) {
            console.error('加载文件失败:', error);
            alert('加载文件失败，请检查文件格式');
        }
    });
    
    // 合并Excel后转换按钮事件
    mergeExcelBtn.addEventListener('click', async () => {
        if (selectedFiles.length < 2) {
            alert('合并功能需要至少2个Excel文件！');
            return;
        }
        
        try {
            mergeExcelBtn.disabled = true;
            mergeExcelBtn.textContent = '合并中...';
            
            const merged = await mergeExcelFiles(selectedFiles);
            mergedWorkbook = merged;
            
            // 显示列选择器让用户选择要保留的列
            showColumnSelector(merged);
            
            mergeExcelBtn.disabled = false;
            mergeExcelBtn.textContent = '合并Excel后转换';
        } catch (error) {
            console.error('合并文件失败:', error);
            alert('合并文件失败：' + error.message);
            mergeExcelBtn.disabled = false;
            mergeExcelBtn.textContent = '合并Excel后转换';
        }
    });
    
    // 合并Excel后转换为CSV按钮事件
    mergeToCsvBtn.addEventListener('click', async () => {
        if (selectedFiles.length < 2) {
            alert('合并功能需要至少2个Excel文件！');
            return;
        }
        
        try {
            mergeToCsvBtn.disabled = true;
            mergeToCsvBtn.textContent = '合并转换中...';
            
            const mergedWorkbook = await mergeExcelFiles(selectedFiles);
            
            // 从合并的工作簿中提取数据
            const firstSheetName = mergedWorkbook.SheetNames[0];
            const worksheet = mergedWorkbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            const csvContent = convertToCSV(jsonData);
            downloadCSV(csvContent);
            
            mergeToCsvBtn.disabled = false;
            mergeToCsvBtn.textContent = '合并Excel后转换为CSV';
        } catch (error) {
            console.error('合并转换CSV失败:', error);
            alert('合并转换CSV失败：' + error.message);
            mergeToCsvBtn.disabled = false;
            mergeToCsvBtn.textContent = '合并Excel后转换为CSV';
        }
    });
    
    // 批量转换处理
    convertBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) {
            alert('请先选择Excel文件！');
            return;
        }
        
        // 直接进行转换，不显示列选择器
        await performBatchConversion();
    });
    
    // 执行批量转换
    async function performBatchConversion() {
        convertBtn.disabled = true;
        convertBtn.textContent = '转换中...';
        markdownOutputTextarea.value = '';
        
        let allMarkdownContent = '';
        
        for (let i = 0; i < selectedFiles.length; i++) {
            const file = selectedFiles[i];
            
            if (i > 0) {
                allMarkdownContent += '\n\n---\n\n';
            }
            
            allMarkdownContent += `# ${file.name}\n\n`;
            
            try {
                const fileContent = await processFile(file);
                allMarkdownContent += fileContent;
            } catch (error) {
                console.error(`处理文件 ${file.name} 时出错:`, error);
                allMarkdownContent += `处理文件 ${file.name} 时出错: ${error.message}\n`;
            }
        }
        
        markdownOutputTextarea.value = allMarkdownContent;
        convertBtn.disabled = false;
        convertBtn.textContent = '开始转换';
        updateButtonStates();
    }
    
    // 加载工作簿
    function loadWorkbook(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }
    
    // 合并Excel文件
    async function mergeExcelFiles(files) {
        const workbooks = [];
        
        // 加载所有工作簿
        for (const file of files) {
            const workbook = await loadWorkbook(file);
            workbooks.push({ workbook, filename: file.name });
        }
        
        // 创建新的合并工作簿
        const mergedWorkbook = XLSX.utils.book_new();
        const allData = [];
        let headers = null;
        
        // 处理每个工作簿的第一个工作表
        for (let i = 0; i < workbooks.length; i++) {
            const { workbook, filename } = workbooks[i];
            const firstSheetName = workbook.SheetNames[0];
            
            if (!firstSheetName) continue;
            
            const worksheet = workbook.Sheets[firstSheetName];
            if (!worksheet['!ref']) continue;
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (jsonData.length === 0) continue;
            
            // 第一个文件的表头作为标准
            if (headers === null) {
                headers = jsonData[0];
                // 添加表头到合并数据
                allData.push(headers);
            }
            
            // 添加数据行（跳过表头）
            for (let j = 1; j < jsonData.length; j++) {
                const row = jsonData[j];
                // 确保行长度与表头一致
                while (row.length < headers.length) {
                    row.push('');
                }
                // 可选：添加来源文件信息
                // row.push(filename);
                allData.push(row);
            }
        }
        
        if (allData.length === 0) {
            throw new Error('没有找到可合并的数据');
        }
        
        // 创建合并后的工作表
        const mergedWorksheet = XLSX.utils.aoa_to_sheet(allData);
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedWorksheet, '合并数据');
        
        return mergedWorkbook;
    }
    
    // 处理单个文件
    function processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const result = convertWorkbookToMarkdown(workbook);
                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }
    
    // 显示列选择器
    function showColumnSelector(workbook) {
        const allColumns = new Set();
        
        // 收集所有工作表的列
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet['!ref']) return;
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            if (jsonData.length > 0) {
                const headers = jsonData[0];
                headers.forEach((header, index) => {
                    const columnName = header || `列${index + 1}`;
                    allColumns.add(columnName);
                });
            }
        });
        
        // 清空并填充列列表
        columnList.innerHTML = '';
        selectedColumns.clear();
        
        Array.from(allColumns).forEach(columnName => {
            selectedColumns.add(columnName);
            
            const columnItem = document.createElement('div');
            columnItem.className = 'column-item';
            columnItem.innerHTML = `
                <label>
                    <input type="checkbox" value="${columnName}" checked>
                    <span class="column-name">${columnName}</span>
                    <span class="column-preview">预览数据...</span>
                </label>
            `;
            
            const checkbox = columnItem.querySelector('input[type="checkbox"]');
            checkbox.addEventListener('change', (e) => {
                if (e.target.checked) {
                    selectedColumns.add(columnName);
                } else {
                    selectedColumns.delete(columnName);
                }
                updateSelectAllState();
            });
            
            columnList.appendChild(columnItem);
        });
        
        columnSelector.style.display = 'block';
        updateSelectAllState();
    }
    
    // 更新全选状态
    function updateSelectAllState() {
        const checkboxes = columnList.querySelectorAll('input[type="checkbox"]');
        const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
        
        if (checkedCount === 0) {
            selectAllColumns.indeterminate = false;
            selectAllColumns.checked = false;
        } else if (checkedCount === checkboxes.length) {
            selectAllColumns.indeterminate = false;
            selectAllColumns.checked = true;
        } else {
            selectAllColumns.indeterminate = true;
        }
    }
    
    // 全选/取消全选事件
    selectAllColumns.addEventListener('change', (e) => {
        const checkboxes = columnList.querySelectorAll('input[type="checkbox"]');
        const isChecked = e.target.checked;
        
        selectedColumns.clear();
        checkboxes.forEach(checkbox => {
            checkbox.checked = isChecked;
            if (isChecked) {
                selectedColumns.add(checkbox.value);
            }
        });
    });
    
    // 应用列筛选
    applyColumnFilter.addEventListener('click', async () => {
        if (selectedColumns.size === 0) {
            alert('请至少选择一列！');
            return;
        }
        
        columnSelector.style.display = 'none';
        
        // 使用选定的列进行转换
        const workbookToConvert = mergedWorkbook || currentWorkbookData;
        const filteredContent = convertWorkbookToMarkdown(workbookToConvert, selectedColumns);
        markdownOutputTextarea.value = filteredContent;
        
        // 清理状态
        currentWorkbookData = null;
        mergedWorkbook = null;
        selectedColumns.clear();
        
        updateButtonStates();
    });
    
    // 取消列筛选
    cancelColumnFilter.addEventListener('click', () => {
        columnSelector.style.display = 'none';
        currentWorkbookData = null;
        mergedWorkbook = null;
        selectedColumns.clear();
    });
    
    // 将工作簿转换为Markdown
    function convertWorkbookToMarkdown(workbook, filterColumns = null) {
        let markdownContent = '';

        // 遍历所有工作表
        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];
            
            // 添加工作表标题
            if (index > 0) {
                markdownContent += '\n\n';
            }
            markdownContent += `## ${sheetName}\n\n`;
            
            // 检查工作表是否为空
            if (!worksheet['!ref']) {
                markdownContent += '*此工作表为空*\n';
                return;
            }
            
            // 转换为JSON数组
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (jsonData.length === 0) {
                markdownContent += '*此工作表为空*\n';
                return;
            }
            
            // 处理列筛选
            let processedData = jsonData;
            if (filterColumns && filterColumns.size > 0) {
                const headers = jsonData[0] || [];
                const columnIndices = [];
                
                // 找到需要保留的列索引
                headers.forEach((header, index) => {
                    const columnName = header || `列${index + 1}`;
                    if (filterColumns.has(columnName)) {
                        columnIndices.push(index);
                    }
                });
                
                // 筛选数据
                processedData = jsonData.map(row => {
                    return columnIndices.map(index => row[index] || '');
                });
            }
            
            // 找到最大列数
            const maxCols = Math.max(...processedData.map(row => row.length));
            
            // 填充所有行到相同长度
            const normalizedData = processedData.map(row => {
                const newRow = [...row];
                while (newRow.length < maxCols) {
                    newRow.push('');
                }
                return newRow;
            });
            
            if (normalizedData.length > 0) {
                // 创建表头
                const headers = normalizedData[0].map((cell, index) => cell || `列${index + 1}`);
                markdownContent += '| ' + headers.join(' | ') + ' |\n';
                markdownContent += '|' + headers.map(() => ' --- ').join('|') + '|\n';
                
                // 添加数据行
                for (let i = 1; i < normalizedData.length; i++) {
                    const row = normalizedData[i];
                    markdownContent += '| ' + row.map(cell => (cell || '').toString().replace(/\|/g, '\\|')).join(' | ') + ' |\n';
                }
            }
        });

        return markdownContent;
    }

    copyBtn.addEventListener('click', () => {
        const markdownText = markdownOutputTextarea.value;
        if (markdownText.trim() === '') {
            alert('没有可复制的内容！');
            return;
        }

        navigator.clipboard.writeText(markdownText).then(() => {
            // 临时改变按钮文本以提供反馈
            const originalText = copyBtn.textContent;
            copyBtn.textContent = '已复制！';
            setTimeout(() => {
                copyBtn.textContent = originalText;
            }, 2000);
        }).catch(err => {
            console.error('复制失败:', err);
            alert('复制失败，请手动选择文本复制。');
        });
    });

    exportBtn.addEventListener('click', () => {
        const markdownText = markdownOutputTextarea.value;
        if (markdownText.trim() === '') {
            alert('没有可导出的内容！');
            return;
        }

        // 创建 Blob 对象
        const blob = new Blob([markdownText], { type: 'text/markdown' });
        
        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // 生成文件名
        let filename = 'converted.md';
        if (mergedWorkbook) {
            filename = `合并文件_${new Date().toISOString().slice(0, 10)}.md`;
        } else if (selectedFiles.length === 1) {
            const baseName = selectedFiles[0].name.replace(/\.[^/.]+$/, "");
            filename = baseName + '.md';
        } else if (selectedFiles.length > 1) {
            filename = `批量转换_${selectedFiles.length}个文件.md`;
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        // 临时改变按钮文本以提供反馈
        const originalText = exportBtn.textContent;
        exportBtn.textContent = '已导出！';
        setTimeout(() => {
            exportBtn.textContent = originalText;
        }, 2000);
    });
    
    // 将合并的工作簿数据转换为CSV格式
    function convertToCSV(workbookData) {
        if (!workbookData || workbookData.length === 0) {
            throw new Error('没有数据可转换');
        }
        
        const csvRows = [];
        
        // 处理每一行数据
        workbookData.forEach(row => {
            const csvRow = row.map(cell => {
                // 处理包含逗号、引号或换行符的单元格
                let cellValue = String(cell || '');
                if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
                    // 转义引号并用引号包围
                    cellValue = '"' + cellValue.replace(/"/g, '""') + '"';
                }
                return cellValue;
            });
            csvRows.push(csvRow.join(','));
        });
        
        return csvRows.join('\n');
    }
    
    // 下载CSV文件
    function downloadCSV(csvContent) {
        // 添加BOM以支持中文字符
        const BOM = '\uFEFF';
        const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
        
        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // 生成文件名
        const filename = `合并文件_${new Date().toISOString().slice(0, 10)}.csv`;
        a.download = filename;
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        // 显示成功提示
        alert('CSV文件已成功导出！');
    }
    
    // 初始化按钮状态
    updateButtonStates();
});