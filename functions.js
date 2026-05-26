Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = apply1stLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = apply2ndLevelBullet;
        
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
    await runPowerPointCommand(async (context) => {
        console.log("Setting Zero Margin...");
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
