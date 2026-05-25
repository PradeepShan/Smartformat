// Initialize the add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // Find elements and add click events
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = apply1stLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = apply2ndLevelBullet;
        
        // Add events for space controls
        document.querySelectorAll(".toggle-btn").forEach(btn => {
            btn.addEventListener('click', handleSpaceToggleClick);
        });
        
        console.log("SmartFormat Add-in is ready.");
    }
});

/**
 * --- Migrated Function Placeholders ---
 */

async function runFixBrand() {
    await runPowerPointCommand(async (context) => {
        // Place VBA 'Fix Brand Text/Color' logic here using Office.js
        console.log("Running Fix Brand command...");
    });
}

async function runZeroMargin() {
    await runPowerPointCommand(async (context) => {
        // Place VBA 'Zero Margin' logic here using Office.js
        console.log("Setting Zero Margin...");
    });
}

async function apply1stLevelBullet() {
    await runPowerPointCommand(async (context) => {
        // Place VBA 'Apply 1st Level Bullet' logic here using Office.js
        console.log("Applying 1st Level Bullet...");
    });
}

async function apply2ndLevelBullet() {
    await runPowerPointCommand(async (context) => {
        // Place VBA 'Apply 2nd Level Bullet' logic here using Office.js
        console.log("Applying 2nd Level Bullet...");
    });
}

// Handler for toggle buttons (0, 1x, 2x, etc.)
function handleSpaceToggleClick(event) {
    // 1. Visually update selected state
    const currentGroup = event.target.parentElement;
    currentGroup.querySelector(".toggle-btn.selected").classList.remove("selected");
    event.target.classList.add("selected");
    
    // 2. Identify the type (Before or After)
    const isBefore = currentGroup.id.includes("btnBefore");
    const multiplier = event.target.textContent;
    
    // 3. Trigger PowerPoint logic (Placeholder)
    console.log(`Setting ${isBefore ? "Before" : "After"} space with multiplier: ${multiplier}`);
}

/**
 * --- Office.js Context Wrapper ---
 * Standard boiler plate to handle PowerPoint context execution.
 */
async function runPowerPointCommand(callback) {
    try {
        await PowerPoint.run(callback);
    } catch (error) {
        console.error("Error executing PowerPoint command:", error);
    }
}
