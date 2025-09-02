import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

export interface PDFExportOptions {
  filename?: string;
  orientation?: 'portrait' | 'landscape';
  format?: 'a4' | 'a3' | 'letter';
  quality?: number;
  margin?: number;
  titleText?: string;
  sourceName?: string;
  chartType?: string;
  summary?: { totalRegions: number; totalYears: number; dataPoints: number; averageValue: number; maxValue: number; minValue: number; decimals?: number } | null;
  tooltipTable?: { title?: string; headers: string[]; rows: string[][] } | null;
}

export const exportChartToPDF = async (
  chartElementId: string,
  options: PDFExportOptions = {}
): Promise<void> => {
  try {
    const {
      filename = `chart_${Date.now()}.pdf`,
      orientation = 'landscape',
      format = 'a4',
      quality = 1.0,
      margin = 10
    } = options;

    // Find the chart element
    const chartElement = document.getElementById(chartElementId);
    if (!chartElement) {
      throw new Error(`Chart element with ID '${chartElementId}' not found`);
    }

    // Find the canvas element inside the chart
    const canvasElement = chartElement.querySelector('canvas');
    if (!canvasElement) {
      throw new Error(`Canvas element not found inside chart with ID '${chartElementId}'`);
    }

    // Temporarily show the element if it's hidden
    const originalDisplay = chartElement.style.display;
    if (originalDisplay === 'none') {
      chartElement.style.display = 'block';
    }

    // Capture the chart canvas
    const canvas = await html2canvas(canvasElement, {
      scale: quality,
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      width: canvasElement.width,
      height: canvasElement.height,
    });

    // Restore original display
    if (originalDisplay === 'none') {
      chartElement.style.display = originalDisplay;
    }

    // Create PDF
    const pdf = new jsPDF({
      orientation,
      unit: 'mm',
      format,
    });

    // Get PDF dimensions
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    // Prepare header information
    const lineHeight = 6;
    const dateStr = new Date().toLocaleString();

    // Meta lines (no title)
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(10);
    let metaY = margin + 6;
    if (options.sourceName) {
      pdf.text(`Generated From: ${options.sourceName}`, margin, metaY);
      metaY += lineHeight;
    }
    pdf.text(`Date: ${dateStr}`, margin, metaY);
    metaY += lineHeight;
    if (options.chartType) {
      pdf.text(`Chart Type: ${options.chartType}`, margin, metaY);
      metaY += lineHeight;
    }

    // Calculate image dimensions and position below header block
    const availableHeightForImage = pdfHeight - metaY - lineHeight - margin; // leave space for summary
    const imgWidth = pdfWidth - (margin * 2);
    let imgHeight = (canvas.height * imgWidth) / canvas.width;

    // Add image to PDF
    const imgData = canvas.toDataURL('image/png');

    if (imgHeight > availableHeightForImage) {
      const ratio = availableHeightForImage / imgHeight;
      const scaledWidth = imgWidth * ratio;
      const scaledHeight = imgHeight * ratio;
      pdf.addImage(imgData, 'PNG', margin, metaY, scaledWidth, scaledHeight);
      imgHeight = scaledHeight;
    } else {
      pdf.addImage(imgData, 'PNG', margin, metaY, imgWidth, imgHeight);
    }

    // Data Summary table
    if (options.summary) {
      let tableY = metaY + imgHeight + lineHeight;
      if (tableY + 40 > pdfHeight - margin) {
        pdf.addPage();
        tableY = margin;
      }
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      pdf.text('Data Summary', margin, tableY);
      tableY += 4;
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);

      const decimals = options.summary.decimals ?? 0;
      const rows: Array<[string, string]> = [
        ['Total Regions', String(options.summary.totalRegions)],
        ['Total Years', String(options.summary.totalYears)],
        ['Data Points', String(options.summary.dataPoints)],
        ['Average', (options.summary.averageValue ?? 0).toFixed(decimals)],
        ['Maximum', (options.summary.maxValue ?? 0).toFixed(decimals)],
        ['Minimum', (options.summary.minValue ?? 0).toFixed(decimals)],
      ];

      const col1X = margin;
      const col2X = pdfWidth / 2;
      const rowHeight = 6;
      rows.forEach((r, idx) => {
        const y = tableY + rowHeight * (idx + 1);
        pdf.text(r[0], col1X, y);
        pdf.text(r[1], col2X, y);
      });
    }

    // Add metadata
    pdf.setProperties({
      title: options.titleText || 'Data Visualization Chart',
      subject: 'Chart Export',
      author: 'BPS Data Visualization',
      creator: 'BPS Chart Dashboard'
    });

    // Save the PDF
    pdf.save(filename);

    console.log(`✅ Chart exported to PDF: ${filename}`);
  } catch (error) {
    console.error('❌ PDF export failed:', error);
    throw error;
  }
};

export const exportMultipleChartsToPDF = async (
  chartElementIds: string[],
  options: PDFExportOptions = {}
): Promise<void> => {
  try {
    const {
      filename = `charts_${Date.now()}.pdf`,
      orientation = 'portrait',
      format = 'a4',
      quality = 1.0,
      margin = 10
    } = options;

    const pdf = new jsPDF({
      orientation,
      unit: 'mm',
      format,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    for (let i = 0; i < chartElementIds.length; i++) {
      const chartElementId = chartElementIds[i];
      const chartElement = document.getElementById(chartElementId);

      if (!chartElement) {
        console.warn(`⚠️ Chart element with ID '${chartElementId}' not found, skipping...`);
        continue;
      }

      // Find the canvas element inside the chart
      const canvasElement = chartElement.querySelector('canvas');
      if (!canvasElement) {
        console.warn(`⚠️ Canvas element not found inside chart with ID '${chartElementId}', skipping...`);
        continue;
      }

      // Add new page for each chart except the first one
      if (i > 0) {
        pdf.addPage();
      }

      // Temporarily show the element if it's hidden
      const originalDisplay = chartElement.style.display;
      if (originalDisplay === 'none') {
        chartElement.style.display = 'block';
      }

      // Capture the chart canvas
      const canvas = await html2canvas(canvasElement, {
        scale: quality,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        width: canvasElement.width,
        height: canvasElement.height,
      });

      // Restore original display
      if (originalDisplay === 'none') {
        chartElement.style.display = originalDisplay;
      }

      // Calculate image dimensions
      const imgWidth = pdfWidth - (margin * 2);
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      // Add image to PDF
      const imgData = canvas.toDataURL('image/png');

      if (imgHeight > pdfHeight - (margin * 2)) {
        // If image is too tall, scale it down
        const ratio = (pdfHeight - (margin * 2)) / imgHeight;
        const scaledWidth = imgWidth * ratio;
        const scaledHeight = imgHeight * ratio;
        pdf.addImage(imgData, 'PNG', margin, margin, scaledWidth, scaledHeight);
      } else {
        pdf.addImage(imgData, 'PNG', margin, margin, imgWidth, imgHeight);
      }
    }

    // Add metadata
    pdf.setProperties({
      title: 'Data Visualization Charts',
      subject: 'Multiple Charts Export',
      author: 'BPS Data Visualization',
      creator: 'BPS Chart Dashboard'
    });

    // Save the PDF
    pdf.save(filename);

    console.log(`✅ ${chartElementIds.length} charts exported to PDF: ${filename}`);
  } catch (error) {
    console.error('❌ Multi-chart PDF export failed:', error);
    throw error;
  }
};

export const exportDashboardToPDF = async (
  dashboardElementId: string,
  options: PDFExportOptions = {}
): Promise<void> => {
  try {
    const {
      filename = `dashboard_${Date.now()}.pdf`,
      orientation = 'portrait',
      format = 'a4',
      quality = 0.8,
      margin = 5
    } = options;

    const dashboardElement = document.getElementById(dashboardElementId);
    if (!dashboardElement) {
      throw new Error(`Dashboard element with ID '${dashboardElementId}' not found`);
    }

    // Capture the entire dashboard
    const canvas = await html2canvas(dashboardElement, {
      scale: quality,
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      height: dashboardElement.scrollHeight,
      width: dashboardElement.scrollWidth,
    });

    const pdf = new jsPDF({
      orientation,
      unit: 'mm',
      format,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    const imgWidth = pdfWidth - (margin * 2);
    const imgHeight = (canvas.height * imgWidth) / canvas.width;

    let heightRemaining = imgHeight;
    let position = 0;

    // Add image to PDF, splitting across pages if necessary
    while (heightRemaining > 0) {
      const pageHeight = Math.min(heightRemaining, pdfHeight - (margin * 2));

      // Create a canvas for this page
      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvas.width;
      pageCanvas.height = (canvas.height * pageHeight) / imgHeight;

      const ctx = pageCanvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(
          canvas,
          0, (position * canvas.height) / imgHeight,
          canvas.width, pageCanvas.height,
          0, 0,
          canvas.width, pageCanvas.height
        );

        const pageImgData = pageCanvas.toDataURL('image/png');

        if (position > 0) {
          pdf.addPage();
        }

        pdf.addImage(pageImgData, 'PNG', margin, margin, imgWidth, pageHeight);
      }

      heightRemaining -= pageHeight;
      position += pageHeight;
    }

    // Add metadata
    pdf.setProperties({
      title: 'Data Visualization Dashboard',
      subject: 'Dashboard Export',
      author: 'BPS Data Visualization',
      creator: 'BPS Chart Dashboard'
    });

    pdf.save(filename);

    console.log(`✅ Dashboard exported to PDF: ${filename}`);
  } catch (error) {
    console.error('❌ Dashboard PDF export failed:', error);
    throw error;
  }
};
