// Farap_Word/src/taskpane/taskpane.js

Office.onReady((info) => {
  // We only run the code if the host is Word
  if (info.host === Office.HostType.Word) {
    // Set up all the button click event handlers
    document.getElementById("loadTemplateBtn").onclick = loadTemplate;
    document.getElementById("applyTemplateBtn").onclick = applyTemplateToWord;
    document.getElementById("saveTemplateBtn").onclick = saveTemplate;
    document.getElementById("createTemplateBtn").onclick = createNewTemplate;
    document.getElementById("uploadWordContentBtn").onclick = uploadWordContent;
    document.getElementById("saveWordToHtmlBtn").onclick = saveWordToHtml;

    // Load existing templates from the document settings
    fetchTemplates();
  }
});

// Global variable to hold all templates.
let templates = [];
// A temporary variable to hold the header and footer when a template is loaded or uploaded.
let uploadedHeaderFooter = { header: "", footer: "" };

/**
 * Fetches templates from the document's settings and populates the dropdown.
 */
function fetchTemplates() {
  templates = Office.context.document.settings.get("word_templates");

  // If no templates are found, create a default one.
  if (!templates || !Array.isArray(templates)) {
    templates = [
      {
        id: "1",
        name: "قالب پیش‌فرض", // Default Template
        content: {
          body: '<h1>قالب نمونه</h1><p><span style="color: red;">این متن قرمز است</span></p>',
          header: "<p>هدر پیش‌فرض</p>", // Default Header
          footer: "<p>فوتر پیش‌فرض</p>", // Default Footer
        },
      },
    ];
    saveTemplatesToSettings(); // Save the default template
  }

  const templateSelect = document.getElementById("templateSelect");
  templateSelect.innerHTML = '<option value="">یک قالب انتخاب کنید</option>'; // Select a template
  templates.forEach((template) => {
    const option = document.createElement("option");
    option.value = template.id;
    option.textContent = template.name;
    templateSelect.appendChild(option);
  });
}

/**
 * Saves the entire 'templates' array to the Word document's settings.
 */
function saveTemplatesToSettings() {
  Office.context.document.settings.set("word_templates", templates);
  Office.context.document.settings.saveAsync((asyncResult) => {
    const statusEl = document.getElementById("status_word");

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Action failed. Error: " + asyncResult.error.message);
      statusEl.innerHTML = "❌ خطا در ذخیره تنظیمات: " + asyncResult.error.message;
      statusEl.style.color = "red";
    } else {
      console.log("Settings saved successfully");
      statusEl.innerHTML = "✅ تنظیمات با موفقیت ذخیره شد";
      statusEl.style.color = "green";
    }
  });
}

/**
 * Loads the selected template's content into the text area.
 */
function loadTemplate() {
  const statusEl = document.getElementById("status_word");
  const templateId = document.getElementById("templateSelect").value;

  if (!templateId) {
    alert("لطفاً یک قالب انتخاب کنید"); // Please select a template
    statusEl.innerHTML = "⚠️ هیچ قالبی انتخاب نشده است";
    statusEl.style.color = "orange";
    return;
  }

  const template = templates.find((t) => t.id === templateId);
  if (template) {
    document.getElementById("templateContent").value = template.content.body;
    uploadedHeaderFooter = {
      header: template.content.header || "",
      footer: template.content.footer || "",
    };

    statusEl.innerHTML = "✅ قالب با موفقیت بارگذاری شد";
    statusEl.style.color = "green";
  } else {
    alert("قالب یافت نشد");
    statusEl.innerHTML = "❌ قالب انتخاب‌شده یافت نشد";
    statusEl.style.color = "red";
  }
}

/**
 * Applies the content (body, header, footer) from the UI to the current Word document.
 */
async function applyTemplateToWord() {
  const bodyContent = document.getElementById("templateContent").value;
  // Get header/footer from the temporary variable
  const headerContent = uploadedHeaderFooter.header;
  const footerContent = uploadedHeaderFooter.footer;

  try {
    await Word.run(async (context) => {
      // Apply body content
      const body = context.document.body;
      body.clear();
      body.insertHtml(bodyContent, Word.InsertLocation.start);

      // Apply header content
      const primaryHeader = context.document.sections.getFirst().getHeader("Primary");
      primaryHeader.clear();
      primaryHeader.insertHtml(headerContent, Word.InsertLocation.start);

      // Apply footer content
      const primaryFooter = context.document.sections.getFirst().getFooter("Primary");
      primaryFooter.clear();
      primaryFooter.insertHtml(footerContent, Word.InsertLocation.start);

      await context.sync();
      alert("قالب با موفقیت اعمال شد."); // Template applied successfully.
    });
  } catch (error) {
    console.error("Error applying template:", error);
    alert("خطا در اعمال قالب: " + error.message); // Error applying template
  }
}

/**
 * Saves changes to an existing template.
 */
function saveTemplate() {
  const templateId = document.getElementById("templateSelect").value;
  const bodyContent = document.getElementById("templateContent").value;

  if (!templateId) {
    alert("لطفاً برای ویرایش، یک قالب را انتخاب و بارگذاری کنید"); // Please select and load a template to edit
    return;
  }

  const template = templates.find((t) => t.id === templateId);
  if (template) {
    template.content.body = bodyContent;
    template.content.header = uploadedHeaderFooter.header; // Use the stored header
    template.content.footer = uploadedHeaderFooter.footer; // Use the stored footer
    saveTemplatesToSettings();
    alert("قالب با موفقیت ویرایش و ذخیره شد"); // Template successfully edited and saved
  } else {
    alert("قالب یافت نشد"); // Template not found
  }
}

/**
 * Creates a new template from the content in the UI.
 */
function createNewTemplate() {
  const templateName = document.getElementById("newTemplateName").value;
  const bodyContent = document.getElementById("templateContent").value;
  if (!templateName) {
    alert("لطفاً نام قالب را وارد کنید"); // Please enter a name for the template
    return;
  }

  const newTemplate = {
    id: Date.now().toString(),
    name: templateName,
    content: {
      body: bodyContent || "<p>قالب جدید</p>", // New template
      header: uploadedHeaderFooter.header, // Use the uploaded header
      footer: uploadedHeaderFooter.footer, // Use the uploaded footer
    },
  };
  templates.push(newTemplate);
  saveTemplatesToSettings();
  fetchTemplates(); // Refresh the dropdown list

  // Clear the input fields
  document.getElementById("newTemplateName").value = "";
  document.getElementById("templateContent").value = "";
  alert("قالب جدید با موفقیت ایجاد شد"); // New template created successfully
}

/**
 * Uploads the full content (body, header, footer) from the Word document into the add-in.
 */
async function uploadWordContent() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const header = context.document.sections.getFirst().getHeader("Primary");
      const footer = context.document.sections.getFirst().getFooter("Primary");

      const bodyHtmlResult = body.getHtml();
      const headerHtmlResult = header.getHtml();
      const footerHtmlResult = footer.getHtml();
      await context.sync();

      // Put the content into the UI and the temporary variable
      document.getElementById("templateContent").value = bodyHtmlResult.value;
      uploadedHeaderFooter = {
        header: headerHtmlResult.value,
        footer: footerHtmlResult.value,
      };

      alert("محتوای کامل سند (شامل هدر و فوتر) با موفقیت آپلود شد."); // Full document content... uploaded successfully.
    });
  } catch (error) {
    console.error("Error uploading Word content:", error);
    alert("خطا در آپلود محتوای ورد: " + error.message); // Error uploading Word content
  }
}

/**
 * Saves the current document's content as a single HTML file.
 */
async function saveWordToHtml() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const header = context.document.sections.getFirst().getHeader("Primary");
      const footer = context.document.sections.getFirst().getFooter("Primary");

      const bodyHtmlResult = body.getHtml();
      const headerHtmlResult = header.getHtml();
      const footerHtmlResult = footer.getHtml();
      await context.sync();

      // Construct a full HTML document string
      const fullHtmlContent = `
        <!DOCTYPE html>
        <html>
        <head> 
          <title>Word Document</title>
          <style>
            body { font-family: 'Times New Roman', Times, serif; }
            header, footer { border: 1px solid #ccc; padding: 10px; margin-bottom: 20px; }
          </style>
        </head>
        <body>
          <header>
            <h2>Header</h2>
            ${headerHtmlResult.value}
          </header>
          <main>
            <h2>Body</h2>
            ${bodyHtmlResult.value}
          </main>
          <footer>
            <h2>Footer</h2>
            ${footerHtmlResult.value}
          </footer>
        </body>
        </html>`;

      const fileName =
        prompt("نام فایل HTML را وارد کنید:", "word_document.html") || "word_document.html";
      saveToHtmlFile(fileName, fullHtmlContent);
    });
  } catch (error) {
    console.error("Error saving Word content to HTML:", error);
    alert("خطا در ذخیره محتوای ورد: " + error.message); // Error saving Word content
  }
}

/**
 * Helper function to trigger a download of text content as a file.
 * @param {string} fileName The desired name of the file.
 * @param {string} content The text content to save.
 */
function saveToHtmlFile(fileName, content) {
  try {
    const blob = new Blob([content], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = fileName.endsWith(".html") ? fileName : `${fileName}.html`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Error in saveToHtmlFile:", error);
    alert("خطا در ذخیره فایل HTML: " + error.message); // Error saving HTML file
  }
}
