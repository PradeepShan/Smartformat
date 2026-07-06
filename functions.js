// ==========================================
// 1. CONSOLE HIJACK & TAB UI LOGIC
// ==========================================
const uiConsole = document.getElementById('on-screen-console');

function printToUI(message, isError = false) {
    if (!uiConsole) return;
    if (typeof message === 'object') {
        try { message = JSON.stringify(message, null, 2); } catch(e) { message = String(message); }
    }
    const msgDiv = document.createElement('div');
    msgDiv.className = isError ? 'log-err' : 'log-msg';
    msgDiv.textContent = `> ${message}`;
    uiConsole.appendChild(msgDiv);
    uiConsole.scrollTop = uiConsole.scrollHeight;
}

const originalLog = console.log;
const originalError = console.error;

console.log = function(...args) {
    originalLog.apply(console, args);
    printToUI(args.join(' '));
};
console.error = function(...args) {
    originalError.apply(console, args);
    printToUI(args.join(' '), true);
};

// Tab Switcher
function switchTab(event, tabId) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');
    event.currentTarget.classList.add('active');
}

// ==========================================
// 2. OFFICE.JS INITIALIZATION
// ==========================================
console.log("Functions.js loaded successfully.");

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Office.js is ready. Wiring up buttons...");
        
        // Brand & Formatting
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = runFirstLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = runSecondLevelBullet;
        
        // Copy/Paste
        document.getElementById("btnCopy").onclick = runCopy;
        document.getElementById("btnPaste").onclick = runPaste;
        
        // Alignment
        const alignIds = ['btnAlignLeft', 'btnAlignCenterH', 'btnAlignCenterV', 'btnAlignRight', 'btnAlignTop', 'btnAlignBottom', 'btnAlignMiddle'];
        alignIds.forEach(id => { document.getElementById(id).onclick = () => runAlign(id); });

        // Distribute
        const distIds = ['btnDistributeH', 'btnDistributeV', 'btnDistributeSpacingH', 'btnDistributeSpacingV'];
        distIds.forEach(id => { document.getElementById(id).onclick = () => runDistribute(id); });

        // Fallbacks for functions requiring VBA
        const logVBA = (name) => console.log(`The '${name}' function requires VBA because JS cannot read PowerPoint Theme logic or Text Runs.`);
        document.getElementById("btnThemeFont").onclick = () => logVBA("Theme Font");
        document.getElementById("btnThemeColor").onclick = () => logVBA("Theme Color");
        document.getElementById("btnMerge").onclick = () => logVBA("Merge");
        document.querySelectorAll(".space-btn").forEach(btn => btn.onclick = () => logVBA("Paragraph Spacing"));

        // Trigger Icon Fetch on first tab click
        let iconsLoaded = false;
        document.querySelector('button[onclick="switchTab(event, \'tab-icons\')"]').addEventListener('click', () => {
            if (!iconsLoaded) {
                loadIcons();
                iconsLoaded = true;
            }
        });
    }
});

// ==========================================
// 3. CORE FORMATTING FUNCTIONS
// ==========================================

async function runZeroMargin() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/textFrame");
        await context.sync();
        if (shapes.items.length === 0) return console.log("Select a shape.");
        shapes.items.forEach(shape => {
            if (shape.textFrame) {
                shape.textFrame.leftMargin = 0; shape.textFrame.rightMargin = 0;
                shape.textFrame.topMargin = 0; shape.textFrame.bottomMargin = 0;
            }
        });
        await context.sync();
        console.log("Zero margin applied.");
    });
}

async function runFixBrand() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/textFrame/textRange/font");
        await context.sync();
        if (shapes.items.length === 0) return;
        shapes.items.forEach(shape => {
            if (shape.textFrame && shape.textFrame.textRange) {
                const font = shape.textFrame.textRange.font;
                font.name = "Montserrat"; // Replace with your brand font
                font.color = "#0052cc";   // Replace with your brand color
            }
        });
        await context.sync();
        console.log("Brand formatting applied.");
    });
}

async function runFirstLevelBullet() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/textFrame/textRange/paragraphFormat");
        await context.sync();
        shapes.items.forEach(shape => {
            if (shape.textFrame && shape.textFrame.textRange) {
                shape.textFrame.textRange.paragraphFormat.bullet.visible = true;
                shape.textFrame.textRange.paragraphFormat.indentLevel = 0; 
            }
        });
        await context.sync();
    });
}

async function runSecondLevelBullet() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/textFrame/textRange/paragraphFormat");
        await context.sync();
        shapes.items.forEach(shape => {
            if (shape.textFrame && shape.textFrame.textRange) {
                shape.textFrame.textRange.paragraphFormat.bullet.visible = true;
                shape.textFrame.textRange.paragraphFormat.indentLevel = 1; 
            }
        });
        await context.sync();
    });
}

// Copy & Paste Globals
let clipboardFormat = null;
async function runCopy() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/left, items/top, items/width, items/height, items/textFrame/leftMargin, items/textFrame/rightMargin, items/textFrame/topMargin, items/textFrame/bottomMargin, items/textFrame/verticalAlignment, items/textFrame/autoSizeSetting, items/textFrame/wordWrap, items/textFrame/textRange/font");
        await context.sync();
        if (shapes.items.length === 0) return console.log("Select a shape to copy.");
        const shape = shapes.items[0];
        clipboardFormat = { left: shape.left, top: shape.top, width: shape.width, height: shape.height, textFrame: null, font: null };
        if (shape.textFrame) {
            clipboardFormat.textFrame = { leftMargin: shape.textFrame.leftMargin, rightMargin: shape.textFrame.rightMargin, topMargin: shape.textFrame.topMargin, bottomMargin: shape.textFrame.bottomMargin, verticalAlignment: shape.textFrame.verticalAlignment, autoSizeSetting: shape.textFrame.autoSizeSetting, wordWrap: shape.textFrame.wordWrap };
            if (shape.textFrame.textRange && shape.textFrame.textRange.font) {
                const f = shape.textFrame.textRange.font;
                clipboardFormat.font = { name: f.name, size: f.size, color: f.color, bold: f.bold, italic: f.italic };
            }
        }
        console.log("Format copied!");
    });
}

async function runPaste() {
    if (!clipboardFormat) return console.log("Clipboard empty.");
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/textFrame/textRange/font");
        await context.sync();
        shapes.items.forEach(shape => {
            shape.left = clipboardFormat.left; shape.top = clipboardFormat.top;
            shape.width = clipboardFormat.width; shape.height = clipboardFormat.height;
            if (shape.textFrame && clipboardFormat.textFrame) {
                const src = clipboardFormat.textFrame;
                shape.textFrame.leftMargin = src.leftMargin; shape.textFrame.rightMargin = src.rightMargin;
                shape.textFrame.topMargin = src.topMargin; shape.textFrame.bottomMargin = src.bottomMargin;
                shape.textFrame.verticalAlignment = src.verticalAlignment; shape.textFrame.autoSizeSetting = src.autoSizeSetting; shape.textFrame.wordWrap = src.wordWrap;
                if (clipboardFormat.font && shape.textFrame.textRange) {
                    const tf = shape.textFrame.textRange.font; const sf = clipboardFormat.font;
                    tf.name = sf.name; tf.size = sf.size; tf.color = sf.color; tf.bold = sf.bold; tf.italic = sf.italic;
                }
            }
        });
        await context.sync();
        console.log("Format pasted!");
    });
}

async function runAlign(actionId) {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/left, items/top, items/width, items/height");
        await context.sync();
        if (shapes.items.length < 2) return console.log("Select at least 2 shapes.");
        
        const keyShape = shapes.items[shapes.items.length - 1]; // Z-Order Top
        for (let i = 0; i < shapes.items.length - 1; i++) {
            const shape = shapes.items[i];
            if (actionId === 'btnAlignLeft') shape.left = keyShape.left;
            if (actionId === 'btnAlignRight') shape.left = keyShape.left + keyShape.width - shape.width;
            if (actionId === 'btnAlignTop') shape.top = keyShape.top;
            if (actionId === 'btnAlignBottom') shape.top = keyShape.top + keyShape.height - shape.height;
            if (actionId === 'btnAlignCenterV') shape.left = keyShape.left + (keyShape.width / 2) - (shape.width / 2);
            if (actionId === 'btnAlignCenterH') shape.top = keyShape.top + (keyShape.height / 2) - (shape.height / 2);
            if (actionId === 'btnAlignMiddle') {
                shape.left = keyShape.left + (keyShape.width / 2) - (shape.width / 2);
                shape.top = keyShape.top + (keyShape.height / 2) - (shape.height / 2);
            }
        }
        await context.sync();
    });
}

async function runDistribute(actionId) {
    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items/left, items/top, items/width, items/height");
        await context.sync();
        
        let shapes = [...selectedShapes.items];
        const cmToPt = 28.3465;
        const exactGapPt = parseFloat(document.getElementById("numDistributeValue").value) * cmToPt;

        if (actionId.includes('H')) {
            shapes.sort((a, b) => a.left - b.left);
            if (actionId === 'btnDistributeSpacingH') {
                for (let i = 1; i < shapes.length; i++) shapes[i].left = shapes[i-1].left + shapes[i-1].width + exactGapPt;
            } else {
                let tw = 0; shapes.forEach(s => tw += s.width);
                const gap = ((shapes[shapes.length - 1].left + shapes[shapes.length - 1].width) - shapes[0].left - tw) / (shapes.length - 1);
                for (let i = 1; i < shapes.length - 1; i++) shapes[i].left = shapes[i-1].left + shapes[i-1].width + gap;
            }
        } else {
            shapes.sort((a, b) => a.top - b.top);
            if (actionId === 'btnDistributeSpacingV') {
                for (let i = 1; i < shapes.length; i++) shapes[i].top = shapes[i-1].top + shapes[i-1].height + exactGapPt;
            } else {
                let th = 0; shapes.forEach(s => th += s.height);
                const gap = ((shapes[shapes.length - 1].top + shapes[shapes.length - 1].height) - shapes[0].top - th) / (shapes.length - 1);
                for (let i = 1; i < shapes.length - 1; i++) shapes[i].top = shapes[i-1].top + shapes[i-1].height + gap;
            }
        }
        await context.sync();
    });
}

// ==========================================
// 4. ICON LIBRARY FETCH & INSERT
// ==========================================
const ICON_SIZE_PT = 48.19; 

async function loadIcons() {
    // ! CHANGE THESE DETAILS TO YOUR PUBLIC GITHUB REPO !
    const githubUser = "PradeepShan"; 
    const githubRepo = "Smartformat";
    const folderPath = "assets"; 

    const apiUrl = `https://api.github.com/repos/${githubUser}/${githubRepo}/contents/${folderPath}`;

    try {
        console.log("Fetching icon list from GitHub...");
        const response = await fetch(apiUrl);
        const files = await response.json();
        
        const grid = document.getElementById('iconGrid');
        grid.innerHTML = ""; 
        
        files.forEach(file => {
            if (file.name.match(/\.(svg|png|jpg|jpeg)$/i)) {
                const cleanName = file.name.split('.')[0].replace(/_/g, ' ');
                const div = document.createElement('div');
                div.className = 'icon-item';
                div.setAttribute('data-name', cleanName.toLowerCase());
                
                div.innerHTML = `
                    <img src="${file.download_url}" alt="${cleanName}">
                    <div class="icon-name">${cleanName}</div>
                `;
                
               div.onclick = () => insertIcon(file.download_url);
                grid.appendChild(div);
            }
        });
        console.log("Icons loaded.");
    } catch (error) {
        console.error("Icon fetch failed. Is the repo public?", error);
    }
}

function filterIcons() {
    const input = document.getElementById('iconSearchInput').value.toLowerCase();
    document.querySelectorAll('.icon-item').forEach(item => {
        item.style.display = item.getAttribute('data-name').includes(input) ? "block" : "none";
    });
}

async function getBase64Image(url) {
    const response = await fetch(url);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

async function insertIcon(url) {
    console.log("Evaluating slide for Insert vs. Swap...");
    
    try {
        const base64Image = await getBase64Image(url);
        
        // DEFAULT: New insertions drop at 100,100 and use the 1.7 cm (48.19 pt) size
        let pLeft = 100; 
        let pTop = 100;
        let pWidth = 48.19;
        let pHeight = 48.19;
        let isSwap = false;
        
        // STEP 1: The "Smart Check" - Does the user have a shape selected?
        await PowerPoint.run(async (context) => {
            try {
                const shapes = context.presentation.getSelectedShapes();
                shapes.load("items/left, items/top, items/width, items/height");
                
                // If NOTHING is selected, this exact line triggers an error and jumps to the 'catch' block below
                await context.sync(); 
                
                if (shapes.items.length > 0) {
                    isSwap = true;
                    const oldIcon = shapes.items[0];
                    pLeft = oldIcon.left; 
                    pTop = oldIcon.top;
                    pWidth = oldIcon.width;
                    pHeight = oldIcon.height;
                    
                    oldIcon.delete(); 
                    await context.sync();
                }
            } catch (error) {
                // We specifically catch the ItemNotFound error here.
                // It just means the user clicked the slide background. We ignore it and proceed with a fresh insert!
                console.log("Nothing selected. Defaulting to fresh insert.");
            }
        });
        
        console.log(isSwap ? "Executing Swap..." : "Executing Fresh Insert...");
        
        // STEP 2: Use the universally supported Common API for insertion
        Office.context.document.setSelectedDataAsync(
            base64Image,
            { 
                coercionType: Office.CoercionType.Image, 
                imageLeft: pLeft, 
                imageTop: pTop, 
                imageWidth: pWidth,   
                imageHeight: pHeight, 
                imageAltText: "decorative"
            },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("API ERROR: " + asyncResult.error.message);
                } else {
                    console.log("SUCCESS: Icon processed successfully!");
                }
            }
        );
    } catch (error) { 
        console.error("Fatal Error processing icon: " + error.message); 
    }
}
