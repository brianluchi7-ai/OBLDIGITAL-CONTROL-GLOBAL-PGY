window.addEventListener("message", async (event) => {
  if (!event.data || event.data.action !== "capture_dashboard") return;

  const type = event.data.type;
  const canvas = await html2canvas(document.body, { scale: 2 });
  const imgData = canvas.toDataURL("image/png");

  if (type === "pdf") {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('l', 'pt', 'a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imgWidth = pageWidth - 20;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
    pdf.save("Dashboard_Captura.pdf");
  }

  if (type === "ppt") {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addImage({ data: imgData, x: 0.3, y: 0.3, w: 9.3, h: 5.5 });
    pptx.writeFile({ fileName: "Dashboard_Captura.pptx" });
  }
});
