// --- ON-SCREEN CONSOLE OVERRIDE ---
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
// --- END CONSOLE OVERRIDE ---

// Immediate test log to prove the file loaded
console.log("Functions.js version 2 successfully loaded!");

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Office.js is ready. Attaching button events...");
        
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = apply1stLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = apply2ndLevelBullet;
        // Textbox Properties (Copy / Paste)
        document.getElementById("btnCopy").onclick = runCopy;
        document.getElementById("btnPaste").onclick = runPaste;

// Add this line to your Office.onReady block, replacing the old alignIds array:
const alignIds = ['btnAlignLeft', 'btnAlignCenterH', 'btnAlignCenterV', 'btnAlignRight', 'btnAlignTop', 'btnAlignBottom', 'btnAlignMiddle'];
        alignIds.forEach(id => {
            document.getElementById(id).onclick = () => runAlign(id);
        });

        // Distribute Buttons
        const distIds = ['btnDistributeH', 'btnDistributeV', 'btnDistributeSpacingH', 'btnDistributeSpacingV'];
        distIds.forEach(id => {
            document.getElementById(id).onclick = () => runDistribute(id);
        });
        
        // Listeners for all the space buttons (0, 1x, 2x, etc.)
        document.querySelectorAll(".space-btn").forEach(btn => {
            btn.addEventListener('click', handleSpaceButtonClick);
        });

        console.log("SmartFormat Add-in is ready.");
    }
});

async function runFixBrand() {
    await runPowerPointCommand(async (context) => {
        console.log("Running Fix Brand command...");
    });
}

async function runZeroMargin() {
    console.log("Running Zero Margin...");
    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        
        // Load the textFrame properties so we can modify the margins
        selectedShapes.load("items/textFrame");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            console.log("No shapes selected. Please select a textbox.");
            return;
        }

        selectedShapes.items.forEach(shape => {
            if (shape.textFrame) {
                // PowerPoint API margins are measured in points
                shape.textFrame.leftMargin = 0;
                shape.textFrame.rightMargin = 0;
                shape.textFrame.topMargin = 0;
                shape.textFrame.bottomMargin = 0;
            }
        });

        await context.sync();
        console.log("Zero margins applied successfully!");
        
    }).catch(function (error) {
        console.error("Error setting margins: " + error);
    });
}
async function apply1stLevelBullet() {
    await runPowerPointCommand(async (context) => {
        console.log("Applying 1st Level Bullet...");
    });
}

async function apply2ndLevelBullet() {
    await runPowerPointCommand(async (context) => {
        console.log("Applying 2nd Level Bullet...");
    });
}


// --- ALIGNMENT LOGIC ---
async function runAlign(actionId) {
    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items/left, items/top, items/width, items/height");
        await context.sync();

        if (selectedShapes.items.length < 2) {
            console.log("Please select at least 2 shapes to align.");
            return;
        }

        const shapes = selectedShapes.items;
        // Key Object is the front-most shape (Z-Index Front)
        const keyShape = shapes[shapes.length - 1];

        console.log(`Aligning to Key Object. Action: ${actionId}`);

        // Loop through all shapes EXCEPT the key shape and align them
        for (let i = 0; i < shapes.length - 1; i++) {
            const shape = shapes[i];

            switch (actionId) {
                case 'btnAlignLeft':
                    shape.left = keyShape.left;
                    break;
                case 'btnAlignRight':
                    shape.left = keyShape.left + keyShape.width - shape.width;
                    break;
                case 'btnAlignTop':
                    shape.top = keyShape.top;
                    break;
                case 'btnAlignBottom':
                    shape.top = keyShape.top + keyShape.height - shape.height;
                    break;
                case 'btnAlignCenterH': // Centers ONLY Horizontally
                   shape.top = keyShape.top + (keyShape.height / 2) - (shape.height / 2);
                    break;
                case 'btnAlignCenterV': // Centers ONLY Vertically
                     shape.left = keyShape.left + (keyShape.width / 2) - (shape.width / 2);
                    
                    break;
                case 'btnAlignMiddle': // Centers BOTH Horizontally and Vertically
                    shape.left = keyShape.left + (keyShape.width / 2) - (shape.width / 2);
                    shape.top = keyShape.top + (keyShape.height / 2) - (shape.height / 2);
                    break;
            }
        }

        await context.sync();
        console.log("Alignment complete.");
    }).catch(error => console.error(error));
}

// --- DISTRIBUTE LOGIC ---
async function runDistribute(actionId) {
    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items/left, items/top, items/width, items/height");
        await context.sync();

        if (selectedShapes.items.length < 3 && (actionId === 'btnDistributeH' || actionId === 'btnDistributeV')) {
            console.log("Standard Distribute requires at least 3 shapes.");
            return;
        }

        // We must copy the API array into a standard array so we can sort it physically
        let shapes = [...selectedShapes.items];

        // 1 cm = 28.3465 points
        const cmToPt = 28.3465;
        const inputValueCm = parseFloat(document.getElementById("numDistributeValue").value);
        const exactGapPt = inputValueCm * cmToPt;

        console.log(`Running Distribution: ${actionId}`);

        if (actionId === 'btnDistributeH' || actionId === 'btnDistributeSpacingH') {
            // Sort shapes from Left to Right
            shapes.sort((a, b) => a.left - b.left);

            if (actionId === 'btnDistributeSpacingH') {
                // Exact Custom Spacing
                console.log(`Spacing H exactly by ${inputValueCm} cm (${exactGapPt.toFixed(2)} pt)`);
                for (let i = 1; i < shapes.length; i++) {
                    shapes[i].left = shapes[i-1].left + shapes[i-1].width + exactGapPt;
                }
            } else {
                // Standard Even Distribute between outer bounds
                let totalShapeWidth = 0;
                shapes.forEach(s => totalShapeWidth += s.width);
                
                const totalSpan = (shapes[shapes.length - 1].left + shapes[shapes.length - 1].width) - shapes[0].left;
                const emptySpace = totalSpan - totalShapeWidth;
                const gap = emptySpace / (shapes.length - 1);

                for (let i = 1; i < shapes.length - 1; i++) {
                    shapes[i].left = shapes[i-1].left + shapes[i-1].width + gap;
                }
            }
        } 
        else if (actionId === 'btnDistributeV' || actionId === 'btnDistributeSpacingV') {
            // Sort shapes from Top to Bottom
            shapes.sort((a, b) => a.top - b.top);

            if (actionId === 'btnDistributeSpacingV') {
                // Exact Custom Spacing
                console.log(`Spacing V exactly by ${inputValueCm} cm (${exactGapPt.toFixed(2)} pt)`);
                for (let i = 1; i < shapes.length; i++) {
                    shapes[i].top = shapes[i-1].top + shapes[i-1].height + exactGapPt;
                }
            } else {
                // Standard Even Distribute between outer bounds
                let totalShapeHeight = 0;
                shapes.forEach(s => totalShapeHeight += s.height);
                
                const totalSpan = (shapes[shapes.length - 1].top + shapes[shapes.length - 1].height) - shapes[0].top;
                const emptySpace = totalSpan - totalShapeHeight;
                const gap = emptySpace / (shapes.length - 1);

                for (let i = 1; i < shapes.length - 1; i++) {
                    shapes[i].top = shapes[i-1].top + shapes[i-1].height + gap;
                }
            }
        }

        await context.sync();
        console.log("Distribution complete.");
    }).catch(error => console.error(error));

}
// --- EXTENDED COPY & PASTE LOGIC ---

// Global variable to hold the copied formatting in memory
let clipboardShapeFormat = null;

async function runCopy() {
    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        
        // Load ALL the structural properties we need
        selectedShapes.load("items/left, items/top, items/width, items/height, items/textFrame/leftMargin, items/textFrame/rightMargin, items/textFrame/topMargin, items/textFrame/bottomMargin, items/textFrame/verticalAlignment, items/textFrame/autoSizeSetting, items/textFrame/wordWrap, items/textFrame/textRange/font");
        
        await context.sync();

        if (selectedShapes.items.length === 0) {
            console.log("Please select a shape to copy its format.");
            return;
        }

        // We only copy from the first selected shape
        const shape = selectedShapes.items[0];

        // 1. Save layout and exact position
        clipboardShapeFormat = {
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
            textFrame: null,
            font: null
        };

        // 2. Save TextFrame properties (margins, alignment, auto-fit)
        if (shape.textFrame) {
            clipboardShapeFormat.textFrame = {
                leftMargin: shape.textFrame.leftMargin,
                rightMargin: shape.textFrame.rightMargin,
                topMargin: shape.textFrame.topMargin,
                bottomMargin: shape.textFrame.bottomMargin,
                verticalAlignment: shape.textFrame.verticalAlignment,
                autoSizeSetting: shape.textFrame.autoSizeSetting,
                wordWrap: shape.textFrame.wordWrap
            };

            // 3. Save Font properties if it has text
            if (shape.textFrame.textRange && shape.textFrame.textRange.font) {
                const font = shape.textFrame.textRange.font;
                clipboardShapeFormat.font = {
                    name: font.name,
                    size: font.size,
                    color: font.color,
                    bold: font.bold,
                    italic: font.italic
                };
            }
        }

        console.log("Full format (including position and margins) copied!");
    }).catch(error => console.error("Copy Error: " + error));
}

async function runPaste() {
    if (!clipboardShapeFormat) {
        console.log("Clipboard is empty. Copy a shape first.");
        return;
    }

    await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        
        // We only need to load the font here so we can overwrite it
        selectedShapes.load("items/textFrame/textRange/font");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            console.log("Select a target shape to paste formatting.");
            return;
        }

        console.log("Pasting format to selected shape(s)...");

        // Loop through all selected shapes and apply the saved formatting
        selectedShapes.items.forEach(shape => {
            
            // 1. Apply Position & Dimensions
            shape.left = clipboardShapeFormat.left;
            shape.top = clipboardShapeFormat.top;
            shape.width = clipboardShapeFormat.width;
            shape.height = clipboardShapeFormat.height;

            // 2. Apply TextFrame properties
            if (shape.textFrame && clipboardShapeFormat.textFrame) {
                const srcFrame = clipboardShapeFormat.textFrame;
                shape.textFrame.leftMargin = srcFrame.leftMargin;
                shape.textFrame.rightMargin = srcFrame.rightMargin;
                shape.textFrame.topMargin = srcFrame.topMargin;
                shape.textFrame.bottomMargin = srcFrame.bottomMargin;
                shape.textFrame.verticalAlignment = srcFrame.verticalAlignment;
                shape.textFrame.autoSizeSetting = srcFrame.autoSizeSetting;
                shape.textFrame.wordWrap = srcFrame.wordWrap;

                // 3. Apply Font properties
                if (clipboardShapeFormat.font && shape.textFrame.textRange) {
                    const targetFont = shape.textFrame.textRange.font;
                    const srcFont = clipboardShapeFormat.font;
                    
                    targetFont.name = srcFont.name;
                    targetFont.size = srcFont.size;
                    targetFont.color = srcFont.color;
                    targetFont.bold = srcFont.bold;
                    targetFont.italic = srcFont.italic;
                }
            }
        });

        await context.sync();
        console.log("Format pasted successfully!");
    }).catch(error => console.error("Paste Error: " + error));
}

// Handler for the simple space action buttons
// Universal handler for all space buttons (Before & After)
async function handleSpaceButtonClick(event) {
    // 1. Identify if the user clicked a "Before" or "After" button
    const isBefore = event.target.id.includes("btnBefore");
    const spaceType = isBefore ? "Before" : "After";
    
    // 2. Parse the multiplier from the button (e.g., "0" -> 0, "2x" -> 2)
    const buttonText = event.target.textContent;
    const multiplier = buttonText === "0" ? 0 : parseFloat(buttonText.replace('x', ''));
    
    // 3. Grab the base value from the corresponding input box
    const inputId = isBefore ? "numBeforeValue" : "numAfterValue";
    const baseValue = parseFloat(document.getElementById(inputId).value);
    
    // 4. Calculate the final space to apply
    const targetSpace = baseValue * multiplier;
    console.log(`Applying ${spaceType} Space: ${baseValue} * ${multiplier} = ${targetSpace}pt`);

    // 5. Execute the PowerPoint formatting
    await PowerPoint.run(async (context) => {
        // Get whatever the user currently has selected on the slide
        const selectedShapes = context.presentation.getSelectedShapes();
        
        // We must 'load' the specific properties we want to edit before syncing
        selectedShapes.load("items/textFrame/textRange/paragraphFormat");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            console.log("No shapes selected. Please select a textbox.");
            return;
        }

        // Loop through all selected shapes and apply the spacing
        selectedShapes.items.forEach(shape => {
            // Check if the shape actually contains text
            if (shape.textFrame) {
                const paragraphFormat = shape.textFrame.textRange.paragraphFormat;
                
                if (isBefore) {
                    paragraphFormat.spaceBefore = targetSpace;
                } else {
                    paragraphFormat.spaceAfter = targetSpace;
                }
            }
        });

        // Sync the changes back to PowerPoint
        await context.sync();
        
    }).catch(function (error) {
        console.error("Error applying paragraph formatting: " + error);
    });
}

async function runPowerPointCommand(callback) {
    try {
        await PowerPoint.run(callback);
    } catch (error) {
        console.error("Error executing PowerPoint command:", error);
    }
}
// --- TAB SWITCHING LOGIC ---
function switchTab(event, tabId) {
    // Hide all tab content
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Remove the blue highlight from all buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show the clicked tab and highlight its button
    document.getElementById(tabId).classList.add('active');
    event.currentTarget.classList.add('active');
}
