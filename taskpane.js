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
    const skippedSlides = [];

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      output += `Found ${slides.items.length} slide(s).\n\n`;

      for (let i = 0; i < slides.items.length; i++) {
        try {
          const slide = slides.items[i];
          let shapes = null;
          let layoutShapes = null;

          // --- Load just the font paths; older APIs choke on "type"/"shapeType"
          try {
            shapes = slide.shapes;
            if (shapes) shapes.load("items/textFrame/textRange/font/name");
          } catch (shapesErr) {
            console.warn(`Slide ${i + 1}: couldn't load shapes (${shapesErr.message})`);
            shapes = null;
          }

          try {
            const layout = slide.layout;
            layoutShapes = layout && layout.shapes ? layout.shapes : null;
            if (layoutShapes) layoutShapes.load("items/textFrame/textRange/font/name");
          } catch (layoutErr) {
            console.warn(`Slide ${i + 1}: couldn't load layout shapes (${layoutErr.message})`);
            layoutShapes = null;
          }

          await context.sync();

          const fonts = new Set();
          const layoutFonts = new Set();

          // --- Slide shapes
          if (shapes && shapes.items) {
            for (const shape of shapes.items) {
              try {
                // Pictures won't have a textFrame, so this guard also "skips images"
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
              } catch (shapeErr) {
                console.warn(`Slide ${i + 1}: shape skipped (${shapeErr.message})`);
              }
            }
          }

          // --- Layout (master) shapes
          if (layoutShapes && layoutShapes.items) {
            for (const shape of layoutShapes.items) {
              try {
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
              } catch (shapeErr) {
                console.warn(`Slide ${i + 1}: layout shape skipped (${shapeErr.message})`);
              }
            }
          }

          // --- Check slide fonts
          for (const font of fonts) {
            if (!isFontInstalled(font)) {
              if (!missingFonts[font]) missingFonts[font] = [];
              missingFonts[font].push(i + 1);
            }
          }

          // --- Check layout-only fonts
          for (const font of layoutFonts) {
            const isMissing = !isFontInstalled(font);
            const isUsedInSlide = usedSlideFonts.has(font);
            if (isMissing && !isUsedInSlide) {
              fontsMissingInMaster.add(font);
              if (!missingFonts[font]) missingFonts[font] = [];
              if (!missingFonts[font].includes("Master only")) {
                missingFonts[font].push("Master only");
              }
            }
          }
        } catch (err) {
          const msg = (err && err.message) || "";
          if (/not accessible/i.test(msg) || err?.code === "AccessDenied") {
            skippedSlides.push(i + 1);
            console.error(`Slide ${i + 1}: Skipped - not accessible by Office add-ins (${msg})`);
          } else {
            console.warn(`Slide ${i + 1}: non-fatal error, continuing. (${msg})`);
          }
        }
      }

      // --- Build output
      output = "";

      if (usedSlideFonts.size > 0) {
        output += "=== FONTS USED IN SLIDES ===\n";
        output += [...usedSlideFonts].sort().join(", ") + "\n\n";
      }

      if (Object.keys(missingFonts).length > 0) {
        output += "=== MISSING FONTS ===\n";
        for (const font in missingFonts) {
          const where = missingFonts[font].join(", ");
          const masterNote = usedMasterFonts.has(font) ? " (Master Slides)" : "";
          output += `❌ ${font}  (Slides: ${where})${masterNote}\n`;
        }
        output += "\n";
      } else {
        output += "✅ No missing fonts detected.\n\n";
      }

      if (skippedSlides.length > 0) {
        output += "=== SKIPPED SLIDES ===\n";
        output += "(not accessible by Office add-ins)\n";
        output += skippedSlides.join(", ") + "\n";
      }
    });

    // === OUTPUT TO TASKPANE HTML ===
    const outputElement = document.getElementById("output");
    const copyBtn = document.getElementById("copy-btn");

    outputElement.innerHTML = "";
    copyBtn.style.display = "none";

    const missingFontList = [];

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
