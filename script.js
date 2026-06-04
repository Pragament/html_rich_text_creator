import * as docx from 'docx';

const editor = document.getElementById('richEditor');
const tocContainer = document.getElementById('tocContainer');
const headingSelect = document.getElementById('headingSelect');
const templateSelect = document.getElementById('templateSelect');
const selectionMenu = document.getElementById('selectionMenu');
const tableMenu = document.getElementById('tableMenu');
let savedSelectionRange = null;
let longPressTimer = null;

const documentTemplates = {
    'question-paper': {
        label: 'Question Paper',
        html: `
            <h1>Question Paper</h1>
            <p><strong>Subject:</strong> ____________________ &nbsp; <strong>Class:</strong> ______ &nbsp; <strong>Time:</strong> ______</p>
            <p><strong>Maximum Marks:</strong> ______ &nbsp; <strong>Date:</strong> ____________________</p>
            <h2>Instructions</h2>
            <ol>
                <li>Answer all questions.</li>
                <li>Write answers clearly and show all necessary steps.</li>
                <li>Marks for each question are shown in brackets.</li>
            </ol>
            <h2>Section A: Short Answer Questions</h2>
            <ol>
                <li>Question 1. [2 marks]</li>
                <li>Question 2. [2 marks]</li>
                <li>Question 3. [2 marks]</li>
            </ol>
            <h2>Section B: Long Answer Questions</h2>
            <ol>
                <li>Question 1. [5 marks]</li>
                <li>Question 2. [5 marks]</li>
            </ol>
            <h2>Section C: Application Based Questions</h2>
            <ol>
                <li>Question 1. [10 marks]</li>
            </ol>
        `
    },
    'question-paper-toc': {
        label: 'Question Paper TOC',
        html: `
            <h1>Subject1</h1>
            <h2>Question1</h2>
            <h3>optionA</h3>
            <h3>optionB</h3>
            <h3>optionC</h3>
            <h3>optionD</h3>
            <h2>Question2</h2>
            <h3>optionA</h3>
            <h3>optionB</h3>
            <h3>optionC</h3>
            <h3>optionD</h3>
            <h1>Subject2</h1>
            <h2>Question1</h2>
            <h3>optionA</h3>
            <h3>optionB</h3>
            <h3>optionC</h3>
            <h3>optionD</h3>
            <h2>Question2</h2>
            <h3>optionA</h3>
            <h3>optionB</h3>
            <h3>optionC</h3>
            <h3>optionD</h3>
        `
    },
    'answer-key': {
        label: 'Answer Key',
        html: `
            <h1>Answer Key</h1>
            <p><strong>Subject:</strong> ____________________ &nbsp; <strong>Class:</strong> ______</p>
            <h2>Section A</h2>
            <ol>
                <li>Answer:</li>
                <li>Answer:</li>
                <li>Answer:</li>
            </ol>
            <h2>Section B</h2>
            <ol>
                <li>Answer with marking scheme:</li>
                <li>Answer with marking scheme:</li>
            </ol>
            <h2>Remarks</h2>
            <p>Evaluation notes:</p>
        `
    },
    worksheet: {
        label: 'Worksheet',
        html: `
            <h1>Worksheet</h1>
            <p><strong>Topic:</strong> ____________________ &nbsp; <strong>Name:</strong> ____________________</p>
            <h2>Learning Goals</h2>
            <ul>
                <li>Goal 1</li>
                <li>Goal 2</li>
            </ul>
            <h2>Practice Questions</h2>
            <ol>
                <li>Question 1</li>
                <li>Question 2</li>
                <li>Question 3</li>
            </ol>
            <h2>Challenge</h2>
            <p>Extension task:</p>
        `
    },
    'lesson-plan': {
        label: 'Lesson Plan',
        html: `
            <h1>Lesson Plan</h1>
            <p><strong>Topic:</strong> ____________________ &nbsp; <strong>Duration:</strong> ______</p>
            <h2>Objectives</h2>
            <ul>
                <li>Objective 1</li>
                <li>Objective 2</li>
            </ul>
            <h2>Materials</h2>
            <ul>
                <li>Material 1</li>
                <li>Material 2</li>
            </ul>
            <h2>Lesson Flow</h2>
            <h3>Introduction</h3>
            <p>Starter activity:</p>
            <h3>Main Activity</h3>
            <p>Teaching and practice plan:</p>
            <h3>Assessment</h3>
            <p>Exit ticket or checking strategy:</p>
        `
    },
    blank: {
        label: 'Blank Document',
        html: '<p><br></p>'
    }
};

function getCurrentBlockElement() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return null;
    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    while (node && node !== editor) {
        const tag = node.tagName?.toLowerCase();
        if (['h1','h2','h3','h4','p','div','section','article'].includes(tag)) return node;
        node = node.parentElement;
    }
    return null;
}

function updateHeadingDropdownFromSelection() {
    const block = getCurrentBlockElement();
    if (!block) return;
    const tag = block.tagName.toLowerCase();
    if (tag === 'h1') headingSelect.value = 'H1';
    else if (tag === 'h2') headingSelect.value = 'H2';
    else if (tag === 'h3') headingSelect.value = 'H3';
    else if (tag === 'h4') headingSelect.value = 'H4';
    else headingSelect.value = 'p';
}

headingSelect.addEventListener('change', (e) => {
    const val = e.target.value;
    if (val === 'p') {
        document.execCommand('formatBlock', false, '<p>');
    } else {
        document.execCommand('formatBlock', false, `<${val}>`);
    }
    editor.focus();
    renderTOC();
    updateHeadingDropdownFromSelection();
});

templateSelect.addEventListener('change', (e) => {
    const template = documentTemplates[e.target.value];
    if (!template) return;
    editor.innerHTML = template.html.trim();
    collapseState = {};
    editor.focus();
    renderTOC();
    updateHeadingDropdownFromSelection();
    updateStatus(`${template.label} loaded`, false);
    templateSelect.value = '';
});

function serializeImages(images) {
    return images.map((img, imageIndex) => {
        const src = img.getAttribute('src') || '';
        return {
            index: imageIndex,
            alt: img.getAttribute('alt') || '',
            src,
        };
    });
}

function serializeTables(tables) {
    return tables.map((table, tableIndex) => {
        const rows = table.rows.length;
        const cols = table.rows[0] ? table.rows[0].cells.length : 0;
        const data = Array.from(table.rows).map((row) => {
            return Array.from(row.cells).map((cell) => ({
                html: cell.innerHTML.trim(),
                isHeader: cell.tagName.toLowerCase() === 'th',
            }));
        });
        return {
            index: tableIndex,
            rowCount: rows,
            colCount: cols,
            data,
        };
    });
}

function getHeadingImages(heading) {
    const images = [...heading.querySelectorAll('img')];
    let node = heading.nextElementSibling;
    while (node && !/^H[1-4]$/.test(node.tagName)) {
        images.push(...node.querySelectorAll('img'));
        node = node.nextElementSibling;
    }
    return serializeImages(images);
}

function getHeadingTables(heading) {
    const tables = [...heading.querySelectorAll('table')];
    let node = heading.nextElementSibling;
    while (node && !/^H[1-4]$/.test(node.tagName)) {
        if (node.tagName.toLowerCase() === 'table') tables.push(node);
        tables.push(...node.querySelectorAll('table'));
        node = node.nextElementSibling;
    }
    return serializeTables(tables);
}

function buildTOCNodes() {
    const headings = editor.querySelectorAll('h1, h2, h3, h4');
    const nodes = [];
    const stack = [{ level: 0, children: nodes, parent: null }];
    headings.forEach((heading, idx) => {
        const tag = heading.tagName.toLowerCase();
        let level = 1;
        if (tag === 'h1') level = 1;
        else if (tag === 'h2') level = 2;
        else if (tag === 'h3') level = 3;
        else if (tag === 'h4') level = 4;
        const text = heading.innerText.trim() || `Heading ${idx+1}`;
        if (!heading.id) heading.id = `heading-${Date.now()}-${idx}-${Math.random()}`;
        const node = {
            id: heading.id,
            text,
            level,
            images: getHeadingImages(heading),
            tables: getHeadingTables(heading),
            children: [],
        };
        while (stack.length > 1 && stack[stack.length-1].level >= level) stack.pop();
        const parent = stack[stack.length-1];
        parent.children.push(node);
        stack.push({ level, children: node.children, parent: node });
    });
    return nodes;
}

let collapseState = {};

function renderTOC() {
    const tree = buildTOCNodes();
    if (tree.length === 0) {
        tocContainer.innerHTML = '<div style="color:#94a3b8; text-align:center; padding:20px;">No headings yet</div>';
        return;
    }

    function getImageSrcLabel(src, index) {
        if (src.startsWith('data:')) {
            const mimeMatch = src.match(/^data:([^;]+)/);
            const mime = mimeMatch ? mimeMatch[1] : 'image';
            return `Image ${index + 1}: ${mime} data URL`;
        }
        return src.length > 90 ? `${src.slice(0, 87)}...` : src;
    }

    function renderNode(node) {
        const nodeKey = node.id;
        const isCollapsed = collapseState[nodeKey] === true;
        const hasChildren = node.children && node.children.length > 0;
        const wrapper = document.createElement('div');
        wrapper.className = 'toc-node';
        const itemDiv = document.createElement('div');
        itemDiv.className = 'toc-item-wrapper';
        const toggleSpan = document.createElement('span');
        toggleSpan.className = 'toggle-icon';
        if (hasChildren) {
            toggleSpan.innerHTML = isCollapsed ? '<i class="fas fa-chevron-right"></i>' : '<i class="fas fa-chevron-down"></i>';
            toggleSpan.style.cursor = 'pointer';
            toggleSpan.addEventListener('click', (e) => {
                e.stopPropagation();
                collapseState[nodeKey] = !isCollapsed;
                renderTOC();
            });
        } else {
            toggleSpan.innerHTML = '•';
            toggleSpan.style.opacity = '0.5';
        }
        const titleSpan = document.createElement('span');
        titleSpan.className = `toc-title ${node.level === 1 ? 'h1' : (node.level === 2 ? 'h2' : (node.level === 3 ? 'h3' : 'h4'))}`;
        titleSpan.innerText = node.text;
        titleSpan.style.cursor = 'pointer';
        titleSpan.addEventListener('click', () => {
            const targetElem = document.getElementById(node.id);
            if (targetElem) {
                targetElem.scrollIntoView({ behavior: 'smooth', block: 'start' });
                editor.focus();
            }
        });
        itemDiv.appendChild(toggleSpan);
        itemDiv.appendChild(titleSpan);
        wrapper.appendChild(itemDiv);
        if (node.images.length) {
            node.images.forEach((image) => {
                if (!image.src) return;
                const imageWrapper = document.createElement('span');
                imageWrapper.className = 'toc-image-src';
                const preview = document.createElement('img');
                preview.className = 'toc-image-preview';
                preview.src = image.src;
                preview.alt = image.alt || `Image ${image.index + 1}`;
                preview.title = image.src.startsWith('data:') ? 'Embedded image preview' : image.src;
                preview.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const targetElem = document.getElementById(node.id);
                    if (targetElem) {
                        targetElem.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        editor.focus();
                    }
                });
                imageWrapper.appendChild(preview);
                wrapper.appendChild(imageWrapper);
            });
        }
        if (node.tables && node.tables.length) {
            node.tables.forEach((table) => {
                const tableHeader = document.createElement('div');
                tableHeader.className = 'toc-table-label';
                tableHeader.innerText = `Table ${table.index + 1}: ${table.rowCount}×${table.colCount}`;
                wrapper.appendChild(tableHeader);

                const tableContainer = document.createElement('div');
                tableContainer.className = 'toc-table-container';
                const preview = document.createElement('table');
                preview.className = 'toc-table-preview';

                table.data.forEach((rowData) => {
                    const tr = document.createElement('tr');
                    rowData.forEach((cellData) => {
                        const cell = document.createElement(cellData.isHeader ? 'th' : 'td');
                        cell.innerHTML = cellData.html || '&nbsp;';
                        tr.appendChild(cell);
                    });
                    preview.appendChild(tr);
                });

                tableContainer.appendChild(preview);
                wrapper.appendChild(tableContainer);
            });
        }
        if (hasChildren) {
            const childrenDiv = document.createElement('div');
            childrenDiv.className = `children-container ${isCollapsed ? 'collapsed' : ''}`;
            node.children.forEach(child => childrenDiv.appendChild(renderNode(child)));
            wrapper.appendChild(childrenDiv);
        }
        return wrapper;
    }

    const container = document.createElement('div');
    tree.forEach(rootNode => container.appendChild(renderNode(rootNode)));
    tocContainer.innerHTML = '';
    tocContainer.appendChild(container);
}

function downloadBlob(blob, filename) {
    if (typeof saveAs === 'function') {
        saveAs(blob, filename);
        return;
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function safeTimestamp() {
    return new Date().toISOString().slice(0, 19).replace(/:/g, '-');
}

function exportTOCAsJSON() {
    const tree = buildTOCNodes();
    const jsonStr = JSON.stringify(tree, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    downloadBlob(blob, `toc_export_${safeTimestamp()}.json`);
    updateStatus('TOC exported as JSON', false);
}

document.getElementById('exportTocJsonBtn').addEventListener('click', exportTOCAsJSON);

const observer = new MutationObserver(() => {
    enhanceImages();
    enhanceTables();
    renderTOC();
});
observer.observe(editor, { childList: true, subtree: true, characterData: true });

document.addEventListener('selectionchange', () => {
    updateHeadingDropdownFromSelection();
    updateActiveStates();
});
editor.addEventListener('keyup', updateHeadingDropdownFromSelection);
editor.addEventListener('mouseup', updateHeadingDropdownFromSelection);

function execCmd(command, value = null) {
    document.execCommand(command, false, value);
    editor.focus();
    renderTOC();
    updateHeadingDropdownFromSelection();
}

function updateActiveStates() {
    document.getElementById('boldBtn').classList.toggle('active', document.queryCommandState('bold'));
    document.getElementById('italicBtn').classList.toggle('active', document.queryCommandState('italic'));
    document.getElementById('underlineBtn').classList.toggle('active', document.queryCommandState('underline'));
}

document.getElementById('boldBtn').addEventListener('click', () => execCmd('bold'));
document.getElementById('italicBtn').addEventListener('click', () => execCmd('italic'));
document.getElementById('underlineBtn').addEventListener('click', () => execCmd('underline'));
document.getElementById('alignLeftBtn').addEventListener('click', () => execCmd('justifyLeft'));
document.getElementById('alignCenterBtn').addEventListener('click', () => execCmd('justifyCenter'));
document.getElementById('alignRightBtn').addEventListener('click', () => execCmd('justifyRight'));
document.getElementById('alignJustifyBtn').addEventListener('click', () => execCmd('justifyFull'));
document.getElementById('unorderedListBtn').addEventListener('click', () => execCmd('insertUnorderedList'));
document.getElementById('orderedListBtn').addEventListener('click', () => execCmd('insertOrderedList'));
document.getElementById('indentBtn').addEventListener('click', () => execCmd('indent'));
document.getElementById('outdentBtn').addEventListener('click', () => execCmd('outdent'));
document.getElementById('undoBtn').addEventListener('click', () => execCmd('undo'));
document.getElementById('redoBtn').addEventListener('click', () => execCmd('redo'));
document.getElementById('fontFamilySelect').addEventListener('change', (e) => execCmd('fontName', e.target.value));
document.getElementById('fontSizeSelect').addEventListener('change', (e) => {
    document.execCommand('styleWithCSS', false, true);
    execCmd('fontSize', e.target.value);
});
document.getElementById('clearFormatBtn').addEventListener('click', () => execCmd('removeFormat'));
document.getElementById('insertTableBtn').addEventListener('click', () => {
    insertTableAtCursor(createTableWrapper(['', '', '', '', '', '', '', '', ''], 3));
    updateStatus('Table inserted', false);
});

const localImageInput = document.getElementById('localImageInput');
document.getElementById('insertLocalImageBtn').addEventListener('click', () => localImageInput.click());
localImageInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file && file.type.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (ev) => insertImageAtCursor(ev.target.result);
        reader.readAsDataURL(file);
    }
    localImageInput.value = '';
});

function insertImageAtCursor(src) {
    const img = document.createElement('img');
    img.style.maxWidth = '100%';
    img.style.borderRadius = '12px';
    img.style.margin = '8px 0';
    img.addEventListener('load', () => {
        if (!img.style.width) img.style.width = `${Math.min(img.naturalWidth || 400, 400)}px`;
        syncImageSizeData(img);
    }, { once: true });
    img.src = src;
    const imageNode = createImageResizer(img);
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
        editor.appendChild(imageNode);
        editor.focus();
        return;
    }
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(imageNode);
    range.setStartAfter(imageNode);
    range.setEndAfter(imageNode);
    selection.removeAllRanges();
    selection.addRange(range);
    editor.focus();
    renderTOC();
}

async function pasteImageFromClipboard() {
    try {
        if (!navigator.clipboard?.read) { updateStatus('Clipboard read not supported', true); return; }
        const items = await navigator.clipboard.read();
        for (const item of items) {
            const imgTypes = item.types.filter(t => t.startsWith('image/'));
            if (imgTypes.length) {
                const blob = await item.getType(imgTypes[0]);
                const reader = new FileReader();
                reader.onload = (e) => { insertImageAtCursor(e.target.result); updateStatus('Image inserted', false); };
                reader.readAsDataURL(blob);
                return;
            }
        }
        updateStatus('No image in clipboard', true);
    } catch(err) { updateStatus('Access denied', true); }
}
document.getElementById('insertImageClipboardBtn').addEventListener('click', pasteImageFromClipboard);

function updateStatus(msg, isErr) {
    const sb = document.getElementById('clipboardStatus');
    sb.innerHTML = `<i class="fas ${isErr ? 'fa-exclamation-triangle' : 'fa-check'}"></i> ${msg}`;
    setTimeout(() => { if (sb.innerHTML.includes(msg)) sb.innerHTML = '<i class="fas fa-check-circle"></i> Ready'; }, 2000);
}

function hideMenus() {
    selectionMenu.style.display = 'none';
    tableMenu.style.display = 'none';
}

function showMenu(menu, x, y, buttons) {
    menu.innerHTML = '';
    buttons.forEach(({ label, action }) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = label;
        btn.addEventListener('click', () => {
            hideMenus();
            action();
        });
        menu.appendChild(btn);
    });
    menu.style.left = `${Math.min(x, window.innerWidth - 220)}px`;
    menu.style.top = `${Math.min(y, window.innerHeight - 160)}px`;
    menu.style.display = 'block';
}

function syncImageSizeData(img) {
    const width = Math.round(img.getBoundingClientRect().width || parseFloat(img.style.width) || img.naturalWidth || 400);
    const height = Math.round(img.getBoundingClientRect().height || img.naturalHeight || 300);
    if (img.dataset.docxWidth !== String(width)) img.dataset.docxWidth = width;
    if (img.dataset.docxHeight !== String(height)) img.dataset.docxHeight = height;
}

function createImageResizer(img) {
    if (img.closest('.image-resizer')) return img.closest('.image-resizer');
    img.dataset.imageEnhanced = 'true';
    const wrapper = document.createElement('span');
    wrapper.className = 'image-resizer';
    wrapper.contentEditable = 'false';
    const handle = document.createElement('span');
    handle.className = 'image-resize-handle';
    handle.title = 'Resize image';
    wrapper.appendChild(img);
    wrapper.appendChild(handle);
    handle.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const startX = e.clientX;
        const startWidth = img.getBoundingClientRect().width || 300;
        const ratio = (img.naturalHeight && img.naturalWidth) ? img.naturalHeight / img.naturalWidth : null;
        function onMove(moveEvent) {
            const nextWidth = Math.max(80, Math.min(editor.clientWidth - 40, startWidth + moveEvent.clientX - startX));
            img.style.width = `${Math.round(nextWidth)}px`;
            img.style.height = ratio ? `${Math.round(nextWidth * ratio)}px` : 'auto';
            syncImageSizeData(img);
        }
        function onUp() {
            document.removeEventListener('pointermove', onMove);
            document.removeEventListener('pointerup', onUp);
            renderTOC();
        }
        document.addEventListener('pointermove', onMove);
        document.addEventListener('pointerup', onUp);
    });
    syncImageSizeData(img);
    return wrapper;
}

function enhanceImages(root = editor) {
    root.querySelectorAll('img:not([data-image-enhanced])').forEach((img) => {
        const parent = img.parentNode;
        const nextSibling = img.nextSibling;
        const wrapper = createImageResizer(img);
        if (!wrapper.parentNode && parent) parent.insertBefore(wrapper, nextSibling);
    });
    root.querySelectorAll('.image-resizer img').forEach(syncImageSizeData);
}

function splitSelectedText(text) {
    return text.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
}

function makeCell(content) {
    const td = document.createElement('td');
    td.innerHTML = content || '<br>';
    return td;
}

function escapeHtml(text) {
    return text.replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

function createTableWrapper(cells, columns) {
    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrapper';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'table-options-btn';
    button.contentEditable = 'false';
    button.title = 'Table options';
    button.innerHTML = '<i class="fas fa-ellipsis-h"></i>';
    const table = document.createElement('table');
    const safeColumns = Math.max(1, columns);
    for (let i = 0; i < cells.length; i += safeColumns) {
        const tr = document.createElement('tr');
        cells.slice(i, i + safeColumns).forEach(cell => tr.appendChild(makeCell(cell)));
        while (tr.children.length < safeColumns) tr.appendChild(makeCell(''));
        table.appendChild(tr);
    }
    wrapper.appendChild(button);
    wrapper.appendChild(table);
    attachTableButton(button, table);
    return wrapper;
}

function insertTableAtCursor(wrapper) {
    const selection = window.getSelection();
    const range = selection && selection.rangeCount ? selection.getRangeAt(0) : null;
    if (range && editor.contains(range.commonAncestorContainer)) {
        range.deleteContents();
        range.insertNode(wrapper);
        range.setStartAfter(wrapper);
        range.setEndAfter(wrapper);
        selection.removeAllRanges();
        selection.addRange(range);
    } else {
        editor.appendChild(wrapper);
    }
    editor.focus();
    renderTOC();
}

function getTableCells(table) {
    return Array.from(table.querySelectorAll('td, th')).map(cell => cell.innerHTML.trim() || '<br>');
}

function replaceTable(table, columns) {
    const cells = getTableCells(table);
    const wrapper = table.closest('.table-wrapper');
    const next = createTableWrapper(cells, columns);
    if (wrapper) wrapper.replaceWith(next);
    else table.replaceWith(next);
    renderTOC();
    updateStatus('Table updated', false);
}

function attachTableButton(button, table) {
    if (button.dataset.tableEnhanced) return;
    button.dataset.tableEnhanced = 'true';
    button.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const cellCount = table.querySelectorAll('td, th').length;
        const buttons = [
            { label: 'Convert to single column', action: () => replaceTable(table, 1) },
            { label: 'Convert to single row', action: () => replaceTable(table, cellCount) },
        ];
        if (cellCount % 2 === 0) buttons.push({ label: `Convert to ${cellCount / 2} columns`, action: () => replaceTable(table, cellCount / 2) });
        showMenu(tableMenu, e.clientX, e.clientY, buttons);
    });
}

function enhanceTables(root = editor) {
    root.querySelectorAll('table').forEach((table) => {
        if (table.closest('.table-wrapper')) return;
        const wrapper = document.createElement('div');
        wrapper.className = 'table-wrapper';
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'table-options-btn';
        button.contentEditable = 'false';
        button.title = 'Table options';
        button.innerHTML = '<i class="fas fa-ellipsis-h"></i>';
        table.parentNode.insertBefore(wrapper, table);
        wrapper.appendChild(button);
        wrapper.appendChild(table);
        attachTableButton(button, table);
    });
    root.querySelectorAll('.table-options-btn').forEach((button) => {
        const table = button.closest('.table-wrapper')?.querySelector('table');
        if (table) attachTableButton(button, table);
    });
}

function restoreSavedSelection() {
    if (!savedSelectionRange) return null;
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(savedSelectionRange);
    return savedSelectionRange;
}

function convertSelectionToTable(columns) {
    const range = restoreSavedSelection();
    if (!range) return;
    const lines = splitSelectedText(range.toString());
    if (!lines.length) return;
    const wrapper = createTableWrapper(lines.map(escapeHtml), columns);
    range.deleteContents();
    range.insertNode(wrapper);
    savedSelectionRange = null;
    editor.focus();
    renderTOC();
    updateStatus('Text converted to table', false);
}

function showSelectionTableMenu(x, y) {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount || selection.isCollapsed || !editor.contains(selection.anchorNode)) return;
    const lines = splitSelectedText(selection.toString());
    if (!lines.length) return;
    savedSelectionRange = selection.getRangeAt(0).cloneRange();
    const buttons = [{ label: `Convert to table (${lines.length} columns)`, action: () => convertSelectionToTable(lines.length) }];
    if (lines.length % 2 === 0) buttons.push({ label: `Convert to table (${lines.length / 2} columns)`, action: () => convertSelectionToTable(lines.length / 2) });
    showMenu(selectionMenu, x, y, buttons);
}

editor.addEventListener('contextmenu', (e) => {
    const selection = window.getSelection();
    if (selection && !selection.isCollapsed && editor.contains(selection.anchorNode)) {
        e.preventDefault();
        showSelectionTableMenu(e.clientX, e.clientY);
    }
});

editor.addEventListener('touchstart', (e) => {
    clearTimeout(longPressTimer);
    longPressTimer = setTimeout(() => {
        const touch = e.touches[0];
        if (touch) showSelectionTableMenu(touch.clientX, touch.clientY);
    }, 650);
});
['touchend', 'touchmove', 'touchcancel'].forEach(eventName => editor.addEventListener(eventName, () => clearTimeout(longPressTimer)));
document.addEventListener('click', (e) => {
    if (!selectionMenu.contains(e.target) && !tableMenu.contains(e.target) && !e.target.closest('.table-options-btn')) hideMenus();
});

editor.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.metaKey) {
        const key = e.key.toLowerCase();
        if (e.shiftKey && (e.code === 'Digit7' || e.key === '7' || e.key === '&')) { e.preventDefault(); execCmd('insertOrderedList'); }
        else if (e.shiftKey && (e.code === 'Digit8' || e.key === '8' || e.key === '*')) { e.preventDefault(); execCmd('insertUnorderedList'); }
        else if (e.shiftKey && key === 'z') { e.preventDefault(); execCmd('redo'); }
        else if (key === 'b') { e.preventDefault(); execCmd('bold'); }
        else if (key === 'i') { e.preventDefault(); execCmd('italic'); }
        else if (key === 'u') { e.preventDefault(); execCmd('underline'); }
        else if (key === 'z') { e.preventDefault(); execCmd('undo'); }
        else if (key === 'y') { e.preventDefault(); execCmd('redo'); }
        else if (e.code === 'Space' || key === ' ') { e.preventDefault(); execCmd('removeFormat'); }
    }
    if (e.key === 'Tab') {
        e.preventDefault();
        if (e.shiftKey) execCmd('outdent');
        else execCmd('indent');
    }
});

document.getElementById('exportDocxBtn').addEventListener('click', async () => {
    const btn = document.getElementById('exportDocxBtn');
    const orig = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Generating...';
    btn.disabled = true;
    try {
        const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType, UnderlineType, HeadingLevel, WidthType } = docx;
        function getAlignment(el) {
            const ta = window.getComputedStyle(el).textAlign;
            if (ta === 'center') return AlignmentType.CENTER;
            if (ta === 'right') return AlignmentType.RIGHT;
            if (ta === 'justify') return AlignmentType.JUSTIFIED;
            return AlignmentType.LEFT;
        }
        function getRunProps(el) {
            const props = {};
            if (!el) return props;
            const cs = window.getComputedStyle(el);
            if (cs.fontWeight === '700' || cs.fontWeight === 'bold') props.bold = true;
            if (cs.fontStyle === 'italic') props.italics = true;
            if (cs.textDecoration.includes('underline')) props.underline = { type: UnderlineType.SINGLE };
            return props;
        }
        function base64ToUint8(base64) {
            const bin = atob(base64.split(',')[1]);
            const arr = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
            return arr;
        }
        function getImageTransformation(img) {
            syncImageSizeData(img);
            const rawWidth = Number(img.dataset.docxWidth) || img.naturalWidth || 400;
            const rawHeight = Number(img.dataset.docxHeight) || img.naturalHeight || 300;
            const width = Math.round(Math.min(rawWidth, 600));
            const height = Math.round(rawHeight * (width / rawWidth));
            return { width, height };
        }
        async function processInlineChildren(node) {
            let runs = [];
            for (let child of node.childNodes) {
                if (child.nodeType === Node.ELEMENT_NODE && ['ul', 'ol'].includes(child.tagName.toLowerCase())) continue;
                if (child.nodeType === Node.ELEMENT_NODE && ['p', 'div'].includes(child.tagName.toLowerCase())) {
                    runs.push(...await processInlineChildren(child));
                    continue;
                }
                const childItems = await processNode(child, getAlignment(node));
                childItems.forEach(item => {
                    if (item instanceof TextRun || item instanceof ImageRun) runs.push(item);
                });
            }
            return runs.length ? runs : [new TextRun('')];
        }
        async function processList(listNode, level = 0) {
            const paragraphs = [];
            const isOrdered = listNode.tagName.toLowerCase() === 'ol';
            const items = Array.from(listNode.children).filter(child => child.tagName?.toLowerCase() === 'li');
            for (let i = 0; i < items.length; i++) {
                const li = items[i];
                const prefix = isOrdered ? `${i + 1}. ` : '• ';
                const children = [new TextRun(prefix), ...await processInlineChildren(li)];
                paragraphs.push(new Paragraph({ children, alignment: getAlignment(li), indent: { left: level * 360 } }));
                for (let child of li.children) {
                    const tag = child.tagName?.toLowerCase();
                    if (tag === 'ul' || tag === 'ol') paragraphs.push(...await processList(child, level + 1));
                }
            }
            return paragraphs;
        }
        async function processTable(tableNode) {
            const rows = [];
            for (let row of tableNode.rows) {
                const cells = [];
                for (let cell of row.cells) {
                    let children = [];
                    for (let child of cell.childNodes) {
                        const items = await processNode(child, getAlignment(cell));
                        items.forEach(item => {
                            if (item instanceof Paragraph) children.push(item);
                            else if (item instanceof TextRun || item instanceof ImageRun) children.push(new Paragraph({ children: [item] }));
                        });
                    }
                    if (!children.length) children = [new Paragraph({ children: [new TextRun(cell.innerText.trim() || '')] })];
                    cells.push(new TableCell({ children }));
                }
                rows.push(new TableRow({ children: cells }));
            }
            return [new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } })];
        }
        async function processNode(node, align = AlignmentType.LEFT) {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent;
                if (text && text.trim()) return [new TextRun({ text, ...getRunProps(node.parentElement) })];
                return [];
            }
            if (node.nodeType === Node.ELEMENT_NODE) {
                const tag = node.tagName.toLowerCase();
                if (tag === 'button') return [];
                if (tag === 'table') return processTable(node);
                if (tag === 'img') {
                    try {
                        let data;
                        if (node.src.startsWith('data:')) data = base64ToUint8(node.src);
                        else { const res = await fetch(node.src); const blob = await res.blob(); data = new Uint8Array(await blob.arrayBuffer()); }
                        return [new ImageRun({ data, transformation: getImageTransformation(node) })];
                    } catch(e) { return [new TextRun('[Image]')]; }
                }
                if (tag === 'br') return [new TextRun({ text: '\n', break: 1 })];
                if (tag === 'ul' || tag === 'ol') return processList(node);
                let childRuns = [];
                for (let child of node.childNodes) childRuns.push(...await processNode(child, getAlignment(node)));
                if (childRuns.some(item => item instanceof Table)) return childRuns;
                const block = ['p','div','h1','h2','h3','h4','li'];
                if (block.includes(tag)) {
                    let heading = null;
                    if (tag === 'h1') heading = HeadingLevel.HEADING_1;
                    else if (tag === 'h2') heading = HeadingLevel.HEADING_2;
                    else if (tag === 'h3') heading = HeadingLevel.HEADING_3;
                    else if (tag === 'h4') heading = HeadingLevel.HEADING_4;
                    if (heading) return [new Paragraph({ children: childRuns.length ? childRuns : [new TextRun('')], heading, alignment: align })];
                    if (!childRuns.length) childRuns = [new TextRun('')];
                    return [new Paragraph({ children: childRuns, alignment: align })];
                }
                return childRuns;
            }
            return [];
        }
        let allParas = [];
        for (let child of editor.childNodes) {
            const res = await processNode(child);
            for (let item of res) if (item instanceof Paragraph || item instanceof Table) allParas.push(item);
            else if (item instanceof TextRun) allParas.push(new Paragraph({ children: [item] }));
        }
        if (!allParas.length) allParas.push(new Paragraph({ children: [new TextRun('Empty')] }));
        const doc = new Document({ sections: [{ properties: {}, children: allParas }] });
        const blob = await Packer.toBlob(doc);
        downloadBlob(blob, `WordPad_${safeTimestamp()}.docx`);
        updateStatus('DOCX ready', false);
    } catch(e) { alert('Export error: ' + e.message); updateStatus('Export failed', true); }
    finally { btn.innerHTML = orig; btn.disabled = false; }
});

enhanceImages();
enhanceTables();
renderTOC();
updateActiveStates();
