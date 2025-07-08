/* global document, Office, PowerPoint */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("✅ Office is ready in PowerPoint");

    const sideload = document.getElementById("sideload-msg");
    if (sideload) sideload.style.display = "none";

    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    const scanBtn = document.getElementById("scan-fonts");
    if (scanBtn) {
      scanBtn.onclick = runFontChecker;
      console.log("✅ Scan Fonts button is wired up");
    } else {
      console.log("❌ Scan Fonts button not found");
    }
  }
});

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => {
    showToast("✅ Copied to clipboard");
  }, (err) => {
    console.error("Clipboard copy failed:", err);
    showToast("❌ Copy failed");
  });
}


function isFontInstalled(fontName) {
  const testString = "mmmmmmmmmmlli";
  const testSize = "72px";

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  context.font = `${testSize} monospace`;
  const baselineWidth = context.measureText(testString).width;

  context.font = `${testSize} '${fontName}', monospace`;
  const testWidth = context.measureText(testString).width;

  return testWidth !== baselineWidth;
}

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.className = "show";

  setTimeout(() => {
    toast.className = toast.className.replace("show", "");
    setTimeout(() => {
      toast.style.display = "none";
    }, 400);
  }, 2000);

  toast.style.display = "block";
}

async function runFontChecker() {
  console.log("🔍 Scan Fonts button clicked");

  try {
    let output = "Scanning...\n";
    document.getElementById("output").textContent = output;

    const missingFonts = {};
    const usedSlideFonts = new Set();
    const usedMasterFonts = new Set();
    const fontsMissingInMaster = new Set();

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      output += `Found ${slides.items.length} slide(s).\n\n`;

      for (let i = 0; i < slides.items.length; i++) {
        try {
          const slide = slides.items[i];
          let shapes = slide.shapes;
          let layout = slide.layout;
          let layoutShapes = layout ? layout.shapes : null;

          if (shapes) shapes.load("items/textFrame/textRange/font/name");
          if (layoutShapes) layoutShapes.load("items/textFrame/textRange/font/name");
          await context.sync();

          const fonts = new Set();
          const layoutFonts = new Set();

          if (shapes && shapes.items) {
            for (const shape of shapes.items) {
              if (
                shape.textFrame &&
                shape.textFrame.textRange &&
                shape.textFrame.textRange.font &&
                shape.textFrame.textRange.font.name
              ) {
                const font = shape.textFrame.textRange.font.name;
                fonts.add(font);
                usedSlideFonts.add(font);
              }
            }
          }

          if (layoutShapes && layoutShapes.items) {
            for (const shape of layoutShapes.items) {
              if (
                shape.textFrame &&
                shape.textFrame.textRange &&
                shape.textFrame.textRange.font &&
                shape.textFrame.textRange.font.name
              ) {
                const font = shape.textFrame.textRange.font.name;
                layoutFonts.add(font);
                usedMasterFonts.add(font);
              }
            }
          }

          const fontList = [...fonts];
          fontList.forEach((font) => {
            const isMissing = !isFontInstalled(font);
            if (isMissing) {
              if (!missingFonts[font]) missingFonts[font] = [];
              missingFonts[font].push(i + 1);
            }
          });

          const layoutFontList = [...layoutFonts];
          layoutFontList.forEach((font) => {
            const isMissing = !isFontInstalled(font);
            const isUsedInSlide = usedSlideFonts.has(font);
            if (isMissing && !isUsedInSlide) {
              fontsMissingInMaster.add(font);
              if (!missingFonts[font]) missingFonts[font] = [];
              if (!missingFonts[font].includes("Master only")) {
                missingFonts[font].push("Master only");
              }
            }
          });
        } catch (err) {
          output += `Slide ${i + 1}: [Skipped due to error: ${err.message}]\n`;
          console.error("Slide loop error:", err);
        }
      }

      // Build the final output text
      output += "\n";
      if (usedSlideFonts.size > 0) {
        output += "=== FONTS USED IN SLIDES ===\n";
        output += [...usedSlideFonts].sort().join(", ") + "\n\n";
      }

      if (Object.keys(missingFonts).length > 0) {
        output += "=== MISSING FONTS ===\n";
        for (const font in missingFonts) {
          const where = missingFonts[font].join(", ");
          const masterNote = usedMasterFonts.has(font) ? " (Master Slides)" : "";
          output += `❌ ${font} (Slides: ${where})${masterNote}\n`;
        }
      } else {
        output += "✅ No missing fonts detected.\n";
      }
    });

    // === OUTPUT TO TASKPANE HTML ===
    const outputElement = document.getElementById("output");
    const copyBtn = document.getElementById("copy-btn");

    outputElement.innerHTML = "";
    copyBtn.style.display = "none";

    let missingFontList = [];

    output.split("\n").forEach((line) => {
      const div = document.createElement("div");

      if (line.startsWith("=== ")) {
        div.textContent = line;
        div.className = "heading";
        outputElement.appendChild(div);
        return;
      }

      if (line.startsWith("✅")) {
        div.textContent = line;
        div.className = "success";
        outputElement.appendChild(div);
        return;
      }

      const match = line.match(/^❌ (.+?) \(Slides:/);
      if (match) {
        const fontName = match[1];
        const span = document.createElement("span");
        span.textContent = fontName;
        span.className = "copyable";
        span.title = "Click to copy";
        span.onclick = () => copyToClipboard(fontName);

        div.appendChild(document.createTextNode("❌ "));
        div.appendChild(span);
        div.appendChild(document.createTextNode("  " + line.slice(3 + fontName.length)));
        div.className = "missing font-line";

        missingFontList.push(fontName);
      } else {
        div.textContent = line;
        div.className = "font-line";
      }

      outputElement.appendChild(div);
    });

    if (missingFontList.length > 0) {
      copyBtn.style.display = "inline-block";
      copyBtn.onclick = () => copyToClipboard(missingFontList.join(", "));
    }

  } catch (error) {
    console.error("❌ Error in runFontChecker:", error);
    document.getElementById("output").textContent =
      "Error: " + (error.message || error.toString());
  }
}
window.runFontChecker = runFontChecker;