// Initialize the add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // Find elements and add click events
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = apply1stLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = apply2ndLevelBullet;
        
        // Add events for space multipliers (1x-4x)
        document.querySelectorAll(".toggle-btn").forEach(btn => {
            btn.addEventListener('click', handleSpaceToggleClick);
        });

        // Add events for independent Zero buttons
        document.getElementById("btnBefore0").onclick = () => handleZeroClick("Before");
        document.getElementById("btnAfter0").onclick = () => handleZeroClick("After");
        
        console.log("SmartFormat Add-in is ready.");
    }
});

/**
 * --- Migrated Function Placeholders ---
 */

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

// Handler for toggle buttons (1x, 2x, etc.)
function handleSpaceToggleClick(event) {
    const currentGroup = event.target.parentElement;
    
    // Remove selected state from siblings in this group
    const currentlySelected = currentGroup.querySelector(".toggle-btn.selected");
    if (currentlySelected) {
        currentlySelected.classList.remove("selected");
    }
    
    // Add selected state to clicked button
    event.target.classList.add("selected");
    
    // Identify the type (Before or After)
    const isBefore = currentGroup.querySelector("button").id.includes("btnBefore");
    const multiplier = event.target.textContent;
    
    console.log(`Setting ${isBefore ? "Before" : "After"} space with multiplier: ${multiplier}`);
}

// Handler for the independent 0 buttons
function handleZeroClick(type) {
    // Clear any active toggles in the corresponding row
    const toggleGroup = document.querySelector(`#btn${type}1x`).parentElement;
    const currentlySelected = toggleGroup.querySelector(".toggle-btn.selected");
    if (currentlySelected) {
        currentlySelected.classList.remove("selected");
    }

    console.log(`Resetting ${type} space to 0`);
}

/**
 * --- Office.js Context Wrapper ---
 */
async function runPowerPointCommand(callback) {
    try {
        await PowerPoint.run(callback);
    } catch (error) {
        console.error("Error executing PowerPoint command:", error);
    }
}
