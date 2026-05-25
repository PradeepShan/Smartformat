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
function handleSpaceButtonClick(event) {
    const currentGroup = event.target.parentElement;
    
    // Check if we are in the Before or After row based on the button ID
    const isBefore = event.target.id.includes("btnBefore");
    const multiplier = event.target.textContent;
    
    // This is where we will hook up your VBA logic
    console.log(`Applying ${isBefore ? "Before" : "After"} space: ${multiplier}`);
}

async function runPowerPointCommand(callback) {
    try {
        await PowerPoint.run(callback);
    } catch (error) {
        console.error("Error executing PowerPoint command:", error);
    }
}
