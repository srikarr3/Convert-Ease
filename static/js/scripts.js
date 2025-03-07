document.addEventListener("DOMContentLoaded", function() {
  const conversionType = document.getElementById('conversion_type');
  const fileInput = document.getElementById('file');

  const conversionAllowed = {
    "jpg_to_pdf": "image/*",
    "pdf_to_jpg": ".pdf",
    "ppt_to_pdf": ".ppt,.pptx",
    "word_to_pdf": ".doc,.docx",
    "excel_to_pdf": ".xls,.xlsx",
    "pdf_ocr": ".pdf",
    "pdf_to_word": ".pdf",
    "pdf_to_ppt": ".pdf"
  };

  conversionType.addEventListener('change', function() {
    const selected = conversionType.value;
    fileInput.accept = conversionAllowed[selected] || "";
  });

  conversionType.dispatchEvent(new Event('change'));
});
