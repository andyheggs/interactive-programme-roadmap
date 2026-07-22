type ElementPdfOptions = {
  element: HTMLElement;
  title: string;
  fileNameSuffix: string;
};

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "programme-report";
}

export async function exportElementPdf({ element, title, fileNameSuffix }: ElementPdfOptions) {
  const [{ default: jsPDF }, { default: html2canvas }] = await Promise.all([import("jspdf"), import("html2canvas")]);
  document.body.classList.add("snapshot-exporting");
  try {
    const canvas = await html2canvas(element, {
      backgroundColor: getComputedStyle(document.documentElement).getPropertyValue("--page").trim() || "#ffffff",
      scale: 2,
      useCORS: true,
      windowWidth: Math.max(1180, element.scrollWidth),
    });

    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const margin = 8;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const contentWidth = pageWidth - margin * 2;
    const contentHeight = pageHeight - margin * 2;
    const imageWidth = contentWidth;
    const imageHeight = (canvas.height * imageWidth) / canvas.width;
    const imageData = canvas.toDataURL("image/png");

    let remainingHeight = imageHeight;
    let y = margin;
    doc.addImage(imageData, "PNG", margin, y, imageWidth, imageHeight);
    remainingHeight -= contentHeight;

    while (remainingHeight > 1) {
      doc.addPage();
      y = margin - (imageHeight - remainingHeight);
      doc.addImage(imageData, "PNG", margin, y, imageWidth, imageHeight);
      remainingHeight -= contentHeight;
    }

    const pageCount = doc.getNumberOfPages();
    for (let page = 1; page <= pageCount; page += 1) {
      doc.setPage(page);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7);
      doc.setTextColor(100, 114, 105);
      doc.text(`Page ${page} of ${pageCount}`, pageWidth - margin, pageHeight - 4, { align: "right" });
    }

    doc.save(`${fileSlug(title)}-${fileSlug(fileNameSuffix)}.pdf`);
  } finally {
    document.body.classList.remove("snapshot-exporting");
  }
}
