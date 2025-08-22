const express = require("express");
const { google } = require("googleapis");
const { parseDocument } = require("htmlparser2");
const htmlToDocx = require("html-to-docx");
const { JSDOM } = require("jsdom");
const { textContent } = require("domutils");
const { selectAll } = require("css-select");
const fs = require("fs");

const path = require("path");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json({ limit: "10mb" }));

function extractSlidesFromHtml(html) {
  const dom = new JSDOM(html);
  const { document } = dom.window;

  const slides = [];
  let lastSlide = null;
  let buffer = [];

  // A title is a <p> where the whole text is exactly the <strong> text
  const isStrongOnlyTitle = (p) => {
    const strong = p.querySelector("strong");
    if (!strong) return false;
    const allText = p.textContent.replace(/\s+/g, " ").trim();
    const strongText = strong.textContent.replace(/\s+/g, " ").trim();
    return allText === strongText;
  };

  const children = Array.from(document.body.children);

  for (const el of children) {
    if (el.tagName === "P" && isStrongOnlyTitle(el)) {
      // We reached a new title: finalize the previous slide with whatever we've buffered
      if (lastSlide) {
        lastSlide.description = buffer.join("\n\n").trim();
      }
      // Start a new slide with this title
      lastSlide = { title: el.textContent.trim(), description: "" };
      slides.push(lastSlide);
      buffer = [];
      continue;
    }

    // Accumulate text for "content before the next title"
    const text = (el.textContent || "").replace(/\s+/g, " ").trim();
    if (text) buffer.push(text);
  }

  // Attach any trailing content to the last slide so nothing is lost
  if (lastSlide && buffer.length) {
    lastSlide.description = buffer.join("\n\n").trim();
  }

  // Clean up any accidental empties
  return slides.filter(s => s.title && typeof s.description === "string");
}

function extractSlidesFromHtml_SlideShow(html) {
  const dom = new JSDOM(html);
  const { document } = dom.window;

  const h2Elements = [...document.querySelectorAll("h2")];
  const slides = [];

  h2Elements.forEach((h2) => {
    const title = h2.textContent.trim();

    let description = "";
    let sibling = h2.nextElementSibling;

    // Collect plain text until the next <h2>
    while (sibling && sibling.tagName !== "H2") {
      description += sibling.textContent + "\n"; // Keep line breaks
      sibling = sibling.nextElementSibling;
    }

    slides.push({
      title,
      description: description.trim()
    });
  });

  return slides;
}



app.post("/upload-doc", async (req, res) => {
  try {
    const { html_base64, access_token, file_name } = req.body;

    if (!html_base64 || !access_token || !file_name) {
      return res.status(400).json({ error: "Missing required fields." });
    }

    // Decode base64 HTML
    const html = Buffer.from(html_base64, "base64").toString("utf8");

    const isArabic = /[\u0590-\u05FF\u0600-\u06FF]/.test(html);

    // Set up OAuth2 client
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token });

    // Convert HTML to DOCX buffer
    const docxBuffer = await htmlToDocx(html);

    // Save DOCX to temporary file
    const tempPath = path.join(__dirname, "temp.docx");
    fs.writeFileSync(tempPath, docxBuffer);

    // Upload to Google Drive
    const drive = google.drive({ version: "v3", auth: oauth2Client });
    const response = await drive.files.create({
      requestBody: {
        name: file_name,
        mimeType: "application/vnd.google-apps.document",
      },
      media: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        body: fs.createReadStream(tempPath),
      },
      fields: "id",
    });

    // Clean up temp file
    fs.unlinkSync(tempPath);

    // Apply RTL or LTR style at Google Docs level
    const dcoument_ID = response.data.id;
    const docs = google.docs({ version: "v1", auth: oauth2Client });

    const doc = await docs.documents.get({ documentId: dcoument_ID });
    const endIndex = doc.data.body.content[doc.data.body.content.length - 1].endIndex;

    await docs.documents.batchUpdate({
      documentId: dcoument_ID,
      requestBody: {
        requests: [
          {
            updateParagraphStyle: {
              range: {
                startIndex: 1, // skip the first index (0 is the start of doc)
                endIndex: endIndex,  // until the end of doc
              },
              paragraphStyle: {
                direction: isArabic ? "RIGHT_TO_LEFT" : "LEFT_TO_RIGHT",
              },
              fields: "direction",
            },
          },
        ],
      },
    });


    

    // Return Google Docs link
    const docId = response.data.id;
    const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
    res.json({ url: docUrl });

  } catch (error) {
    console.error("Upload error:", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.post("/create-styled-sheet", async (req, res) => {
  try {
    const { access_token, html_base64, title = "Styled Sheet" } = req.body;

    if (!access_token || !html_base64) {
      return res.status(400).json({ error: "Missing 'access_token' or 'html_base64'" });
    }

    // Decode base64 HTML
    const html = Buffer.from(html_base64, "base64").toString("utf8");

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token });

    const sheets = google.sheets({ version: "v4", auth });

    const dom = parseDocument(html);
    const table = selectAll("table", dom)[0];
    const rows = selectAll("tr", table);

    const values = rows.map((row) =>
      selectAll("td, th", row).map((cell) =>
        textContent(cell).trim()
      )
    );

    const sheetRes = await sheets.spreadsheets.create({
      requestBody: {
        properties: { title },
      },
    });

    const spreadsheetId = sheetRes.data.spreadsheetId;

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: "Sheet1!A1",
      valueInputOption: "RAW",
      requestBody: { values },
    });

    // Format header row
    const headerStyle = {
      repeatCell: {
        range: {
          sheetId: 0,
          startRowIndex: 0,
          endRowIndex: 1,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.25, green: 0.32, blue: 0.71 },
            horizontalAlignment: "LEFT",
            textFormat: {
              foregroundColor: { red: 1, green: 1, blue: 1 },
              fontSize: 12,
              bold: true,
            },
          },
        },
        fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
      },
    };

    // Format alternating body rows
    const bodyStyles = values.slice(1).map((_, i) => ({
      repeatCell: {
        range: {
          sheetId: 0,
          startRowIndex: i + 1,
          endRowIndex: i + 2,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: i % 2 === 0
              ? { red: 1, green: 1, blue: 1 }
              : { red: 0.98, green: 0.98, blue: 0.98 },
            textFormat: { fontSize: 11 },
          },
        },
        fields: "userEnteredFormat(backgroundColor,textFormat.fontSize)",
      },
    }));

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [headerStyle, ...bodyStyles],
      },
    });

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
    res.json({ url: sheetUrl });

  } catch (err) {
    console.error("Sheet creation error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// app.post("/create-slides", async (req, res) => {
//   try {
//     const { access_token, html_base64, presentationTitle = "New Slides" } = req.body;

//     if (!access_token || !html_base64) {
//       return res.status(400).json({ error: "Missing 'access_token' or 'html_base64'" });
//     }

//     const htmlContent = Buffer.from(html_base64, "base64").toString("utf8");
//     const slidesData = extractSlidesFromHtml(htmlContent);

//     if (!slidesData.length) {
//       return res.status(400).json({ error: "No valid slides found in HTML." });
//     }

//     const auth = new google.auth.OAuth2();
//     auth.setCredentials({ access_token });

//     const slidesApi = google.slides({ version: "v1", auth });

//     // 1. Create presentation
//     const { data: { presentationId } } = await slidesApi.presentations.create({
//       requestBody: { title: presentationTitle },
//     });

//     // 2. Delete default slide
//     const existingSlides = await slidesApi.presentations.get({ presentationId });
//     const defaultSlideId = existingSlides.data.slides[0].objectId;

//     await slidesApi.presentations.batchUpdate({
//       presentationId,
//       requestBody: {
//         requests: [{ deleteObject: { objectId: defaultSlideId } }],
//       },
//     });

//     // 3. Create new slides
//     const slideRequests = slidesData.map((_, index) => ({
//       createSlide: {
//         objectId: `slide_${index + 1}`,
//         slideLayoutReference: { predefinedLayout: "TITLE_AND_BODY" },
//       },
//     }));

//     await slidesApi.presentations.batchUpdate({
//       presentationId,
//       requestBody: { requests: slideRequests },
//     });

//     // 4. Insert text
//     const newSlides = await slidesApi.presentations.get({ presentationId });
//     const textInsertRequests = [];

//     newSlides.data.slides.forEach((slide, i) => {
//       const { title, description } = slidesData[i] || {};

//       const titleElement = slide.pageElements.find(
//         el => el.shape?.placeholder?.type === "TITLE"
//       );
//       const bodyElement = slide.pageElements.find(
//         el => el.shape?.placeholder?.type === "BODY"
//       );

//       if (title && titleElement) {
//         textInsertRequests.push({
//           insertText: {
//             objectId: titleElement.objectId,
//             text: title,
//             insertionIndex: 0,
//           },
//         });
//       }

//       if (description && bodyElement) {
//         textInsertRequests.push({
//           insertText: {
//             objectId: bodyElement.objectId,
//             text: description,
//             insertionIndex: 0,
//           },
//         });
//       }
//     });

//     if (textInsertRequests.length) {
//       await slidesApi.presentations.batchUpdate({
//         presentationId,
//         requestBody: { requests: textInsertRequests },
//       });
//     }

//     const url = `https://docs.google.com/presentation/d/${presentationId}/edit`;
//     res.json({ url });

//   } catch (err) {
//     console.error("Slides API Error:", err.message);
//     res.status(500).json({ error: err.message });
//   }
// });

app.post("/create-slides", async (req, res) => {
  try {
    const { access_token, html_base64, file_name } = req.body;

    if (!access_token || !html_base64) {
      return res.status(400).json({ error: "Missing 'access_token' or 'html_base64'" });
    }

    const htmlContent = Buffer.from(html_base64, "base64").toString("utf8");
    const slidesData = extractSlidesFromHtml(htmlContent);
    if (!slidesData.length) {
      return res.status(400).json({ error: "No valid slides found in HTML." });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token });

    const slidesApi = google.slides({ version: "v1", auth });

    // Step 1: Create presentation
    const presentationTitle = file_name && file_name.trim() ? file_name : "My Presentation";
    const { data: { presentationId } } = await slidesApi.presentations.create({
      requestBody: { title: presentationTitle },
    });

    // Step 2: Delete default slide
    const defaultSlide = await slidesApi.presentations.get({ presentationId });
    const defaultSlideId = defaultSlide.data.slides[0].objectId;
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: [{ deleteObject: { objectId: defaultSlideId } }] }
    });

    const BACKGROUND_IMAGE_URL = 'https://9b05dd864822d678c9fcbed18bf8311c.cdn.bubble.io/f1753546339862x395890190803218200/background.PNG?_gl=1*1bmqctz*_gcl_au*MTgxNzI2MDE2OC4xNzQ2NDU3OTc5*_ga*MjAzNTQ2NTk5LjE2NzcyMzIwNjY.*_ga_BFPVR2DEE2*czE3NTM1MzQ0NTckbzIxOSRnMSR0MTc1MzU0NjAzNiRqNjAkbDAkaDA.';
    const LOGO_IMAGE_URL = 'https://9b05dd864822d678c9fcbed18bf8311c.cdn.bubble.io/f1753458006954x918763125344364000/46c94753-1589-47cd-8efa-cd023be6bd4a.png?_gl=1*zow64b*_gcl_au*MTgxNzI2MDE2OC4xNzQ2NDU3OTc5*_ga*MjAzNTQ2NTk5LjE2NzcyMzIwNjY.*_ga_BFPVR2DEE2*czE3NTM1MzQ0NTckbzIxOSRnMSR0MTc1MzU0NjAzNiRqNjAkbDAkaDA.';

    const slideRequests = [];

    slidesData.forEach((slide, index) => {
      const slideId = `slide_${index + 1}`;

      // Create slide
      slideRequests.push({
        createSlide: {
          objectId: slideId,
          slideLayoutReference: { predefinedLayout: "BLANK" }
        }
      });

      // Set background image
      slideRequests.push({
        createImage: {
          objectId: `bg_${slideId}`,
          url: BACKGROUND_IMAGE_URL,
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 405, unit: "PT" }, width: { magnitude: 720, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 0,
              translateY: 0,
              unit: "PT"
            }
          }
        }
      });

      // Add logo (bottom left)
      slideRequests.push({
        createImage: {
          objectId: `logo_${slideId}`,
          url: LOGO_IMAGE_URL,
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 40, unit: "PT" }, width: { magnitude: 40, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 10,
              translateY: 340,
              unit: "PT"
            }
          }
        }
      });

      // Add footer (bottom right)
      slideRequests.push({
        createShape: {
          objectId: `footer_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 20, unit: "PT" }, width: { magnitude: 200, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 540,
              translateY: 370,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `footer_${slideId}`,
          text: "HelpMeTeach.AI",
          insertionIndex: 0
        }
      });

      // Add title (centered top)
      slideRequests.push({
        createShape: {
          objectId: `title_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 50, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 60,
              translateY: 50,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `title_${slideId}`,
          text: slide.title,
          insertionIndex: 0
        }
      });
      slideRequests.push({
        updateTextStyle: {
          objectId: `title_${slideId}`,
          style: {
            fontSize: { magnitude: 18, unit: "PT" },
            bold: true
          },
          fields: "fontSize,bold"
        }
      });

      // Add description (centered middle)
      slideRequests.push({
        createShape: {
          objectId: `desc_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 80, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 60,
              translateY: 70,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `desc_${slideId}`,
          text: slide.description,
          insertionIndex: 0
        }
      });
      slideRequests.push({
        updateTextStyle: {
          objectId: `desc_${slideId}`,
          style: {
            fontSize: { magnitude: 14, unit: "PT" },
            bold: false
          },
          fields: "fontSize,bold"
        }
      });
    });

    // Execute all requests
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: slideRequests }
    });

    const url = `https://docs.google.com/presentation/d/${presentationId}/edit`;
    res.json({ url });

  } catch (err) {
    console.error("Slides API Error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post("/create-slides-show", async (req, res) => {
  try {
    const { access_token, html_base64, file_name } = req.body;

    if (!access_token || !html_base64) {
      return res.status(400).json({ error: "Missing 'access_token' or 'html_base64'" });
    }

    const htmlContent = Buffer.from(html_base64, "base64").toString("utf8");
    const slidesData = extractSlidesFromHtml_SlideShow(htmlContent);
    if (!slidesData.length) {
      return res.status(400).json({ error: "No valid slides found in HTML." });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token });

    const slidesApi = google.slides({ version: "v1", auth });

    // Step 1: Create presentation
    const presentationTitle = file_name && file_name.trim() ? file_name : "My Presentation";
    const { data: { presentationId } } = await slidesApi.presentations.create({
      requestBody: { title: presentationTitle },
    });

    // Step 2: Delete default slide
    const defaultSlide = await slidesApi.presentations.get({ presentationId });
    const defaultSlideId = defaultSlide.data.slides[0].objectId;
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: [{ deleteObject: { objectId: defaultSlideId } }] }
    });

    const BACKGROUND_IMAGE_URL = 'https://9b05dd864822d678c9fcbed18bf8311c.cdn.bubble.io/f1753546339862x395890190803218200/background.PNG?_gl=1*1bmqctz*_gcl_au*MTgxNzI2MDE2OC4xNzQ2NDU3OTc5*_ga*MjAzNTQ2NTk5LjE2NzcyMzIwNjY.*_ga_BFPVR2DEE2*czE3NTM1MzQ0NTckbzIxOSRnMSR0MTc1MzU0NjAzNiRqNjAkbDAkaDA.';
    const LOGO_IMAGE_URL = 'https://9b05dd864822d678c9fcbed18bf8311c.cdn.bubble.io/f1753458006954x918763125344364000/46c94753-1589-47cd-8efa-cd023be6bd4a.png?_gl=1*zow64b*_gcl_au*MTgxNzI2MDE2OC4xNzQ2NDU3OTc5*_ga*MjAzNTQ2NTk5LjE2NzcyMzIwNjY.*_ga_BFPVR2DEE2*czE3NTM1MzQ0NTckbzIxOSRnMSR0MTc1MzU0NjAzNiRqNjAkbDAkaDA.';

    const slideRequests = [];

    slidesData.forEach((slide, index) => {
      const slideId = `slide_${index + 1}`;

      // Create slide
      slideRequests.push({
        createSlide: {
          objectId: slideId,
          slideLayoutReference: { predefinedLayout: "BLANK" }
        }
      });

      // Set background image
      slideRequests.push({
        createImage: {
          objectId: `bg_${slideId}`,
          url: BACKGROUND_IMAGE_URL,
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 405, unit: "PT" }, width: { magnitude: 720, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 0,
              translateY: 0,
              unit: "PT"
            }
          }
        }
      });

      // Add logo (bottom left)
      slideRequests.push({
        createImage: {
          objectId: `logo_${slideId}`,
          url: LOGO_IMAGE_URL,
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 40, unit: "PT" }, width: { magnitude: 40, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 10,
              translateY: 340,
              unit: "PT"
            }
          }
        }
      });

      // Add footer (bottom right)
      slideRequests.push({
        createShape: {
          objectId: `footer_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 20, unit: "PT" }, width: { magnitude: 200, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 540,
              translateY: 370,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `footer_${slideId}`,
          text: "HelpMeTeach.AI",
          insertionIndex: 0
        }
      });

      // Add title (centered top)
      slideRequests.push({
        createShape: {
          objectId: `title_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 50, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 60,
              translateY: 50,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `title_${slideId}`,
          text: slide.title,
          insertionIndex: 0
        }
      });
      slideRequests.push({
        updateTextStyle: {
          objectId: `title_${slideId}`,
          style: {
            fontSize: { magnitude: 18, unit: "PT" },
            bold: true
          },
          fields: "fontSize,bold"
        }
      });

      // Add description (centered middle)
      slideRequests.push({
        createShape: {
          objectId: `desc_${slideId}`,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: slideId,
            size: { height: { magnitude: 80, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 60,
              translateY: 100,
              unit: "PT"
            }
          }
        }
      });
      slideRequests.push({
        insertText: {
          objectId: `desc_${slideId}`,
          text: slide.description,
          insertionIndex: 0
        }
      });
      slideRequests.push({
        updateTextStyle: {
          objectId: `desc_${slideId}`,
          style: {
            fontSize: { magnitude: 14, unit: "PT" },
            bold: false
          },
          fields: "fontSize,bold"
        }
      });
    });

    // Execute all requests
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: slideRequests }
    });

    const url = `https://docs.google.com/presentation/d/${presentationId}/edit`;
    res.json({ url });

  } catch (err) {
    console.error("Slides API Error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get("/", (req, res) => {
  res.send("✅ Code running");
});

app.listen(3000, () => {
  console.log("🚀 Server running on port 3000");
});
