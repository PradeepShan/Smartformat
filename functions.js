Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnFixBrand").onclick = runFixBrand;
        document.getElementById("btnZeroMargin").onclick = runZeroMargin;
        document.getElementById("btn1stLevelBullet").onclick = apply1stLevelBullet;
        document.getElementById("btn2ndLevelBullet").onclick = apply2ndLevelBullet;
        
        // This single event listener now handles 0, 1x, 2x, 3x, and 4x
        document.querySelectorAll(".toggle-btn").forEach(btn => {
            btn.addEventListener('click', handleSpaceToggleClick);
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

// Universal handler for the toggle buttons (including 0)
function handleSpaceToggleClick(event) {
    const currentGroup = event.target.parentElement;
    
    // Remove selected state from whichever button is currently active in this row
    const currentlySelected = currentGroup.querySelector(".toggle-btn.selected");
    if (currentlySelected) {
        currentlySelected.classList.remove("selected");
    }
    
    // Add selected state to the button you just clicked
    event.target.classList.add("selected");
    
    // Check if we are in the Before or After row
    const isBefore = currentGroup.querySelector("button").id.includes("btnBefore");
    const multiplier = event.target.textContent;
    
    console.log(`Setting ${isBefore ? "Before" : "After"} space to: ${multiplier}`);
}

async function runPowerPointCommand(callback) {
    try {
        await PowerPoint.run(callback);
    } catch (error) {
        console.error("Error executing PowerPoint command:", error);
    }
}
