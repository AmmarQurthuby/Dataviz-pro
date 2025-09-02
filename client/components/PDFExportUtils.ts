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
  layoutMode?: 'grid' | 'list';
  titles?: string[];
  sourceNames?: string[];
  chartTypesArr?: string[];
  summaries?: ({ totalRegions: number; totalYears: number; dataPoints: number; averageValue: number; maxValue: number; minValue: number; decimals?: number } | null)[];
  tooltipTablesArr?: ({ title?: string; headers: string[]; rows: string[][] } | null)[];
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

    // Use the native chart canvas for best quality
    const srcCanvas = canvasElement as HTMLCanvasElement;

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
    const headerY = margin + 6;
    const lineHeight = 6;
    const dateStr = new Date().toLocaleString();

    // Title
    if (options.titleText) {
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(16);
      pdf.text(String(options.titleText), pdfWidth / 2, headerY, { align: 'center' });
    }

    // Meta lines
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(10);
    let metaY = (options.titleText ? headerY + lineHeight : headerY);
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
    const imageY = metaY + 2;
    const availableHeightForImage = pdfHeight - imageY - lineHeight - margin; // leave space for summary
    const imgWidth = pdfWidth - (margin * 2);

    const srcWidth = (srcCanvas as HTMLCanvasElement).width;
    const srcHeight = (srcCanvas as HTMLCanvasElement).height;
    let imgHeight = (srcHeight * imgWidth) / srcWidth;

    // Add image to PDF using high-quality canvas data
    const imgData = (srcCanvas as HTMLCanvasElement).toDataURL('image/png');

    if (imgHeight > availableHeightForImage) {
      const ratio = availableHeightForImage / imgHeight;
      const scaledWidth = imgWidth * ratio;
      const scaledHeight = imgHeight * ratio;
      pdf.addImage(imgData, 'PNG', margin, imageY, scaledWidth, scaledHeight);
      imgHeight = scaledHeight;
    } else {
      pdf.addImage(imgData, 'PNG', margin, imageY, imgWidth, imgHeight);
    }

    // Data Summary table
    let nextSectionY = imageY + imgHeight + lineHeight;
    if (options.summary) {
      let tableY = nextSectionY;
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
        if (y > pdfHeight - margin) {
          pdf.addPage();
          tableY = margin;
        }
        pdf.text(r[0], col1X, tableY + rowHeight * (idx + 1));
        pdf.text(r[1], col2X, tableY + rowHeight * (idx + 1));
      });
      nextSectionY = tableY + rowHeight * rows.length + lineHeight;
    }

    // Tooltip values table
    if (options.tooltipTable && options.tooltipTable.headers && options.tooltipTable.rows && options.tooltipTable.rows.length > 0) {
      let tableY = nextSectionY;
      if (tableY + 20 > pdfHeight - margin) {
        pdf.addPage();
        tableY = margin;
      }
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      pdf.text(options.tooltipTable.title || 'Data Values', margin, tableY);
      tableY += 6;
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);

      const colCount = options.tooltipTable.headers.length;
      const tableWidth = pdfWidth - margin * 2;
      const colWidth = tableWidth / colCount;
      const rowHeight = 6;

      // Header
      options.tooltipTable.headers.forEach((h, i) => {
        const x = margin + i * colWidth;
        pdf.text(String(h), x, tableY);
      });
      tableY += rowHeight;

      // Rows with pagination
      for (let r = 0; r < options.tooltipTable.rows.length; r++) {
        const row = options.tooltipTable.rows[r];
        const y = tableY + rowHeight;
        if (y > pdfHeight - margin) {
          pdf.addPage();
          tableY = margin;
          // re-render header on new page
          options.tooltipTable.headers.forEach((h, i) => {
            const x = margin + i * colWidth;
            pdf.text(String(h), x, tableY);
          });
          tableY += rowHeight;
        }
        row.forEach((cell, i) => {
          const x = margin + i * colWidth;
          pdf.text(String(cell), x, tableY);
        });
        tableY += rowHeight;
      }
    }

    // Add metadata
    pdf.setProperties({
      title: options.titleText || options.sourceName || 'Data Visualization Chart',
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
      quality = 1.5,
      margin = 10
    } = options;

    const pdf = new jsPDF({
      orientation,
      unit: 'mm',
      format,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();
    const lineHeight = 6;

    // Helper to draw a summary table
    const renderSummary = (summary: any, startY: number, title?: string): number => {
      let y = startY;
      if (title) {
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(12);
        pdf.text(title, margin, y);
        y += 4;
      }
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      const decimals = summary?.decimals ?? 0;
      const rows: Array<[string, string]> = [
        ['Total Regions', String(summary?.totalRegions ?? 0)],
        ['Total Years', String(summary?.totalYears ?? 0)],
        ['Data Points', String(summary?.dataPoints ?? 0)],
        ['Average', (summary?.averageValue ?? 0).toFixed(decimals)],
        ['Maximum', (summary?.maxValue ?? 0).toFixed(decimals)],
        ['Minimum', (summary?.minValue ?? 0).toFixed(decimals)],
      ];
      const col1X = margin;
      const col2X = pdfWidth / 2;
      const rowH = 6;
      rows.forEach((r, idx) => {
        const yy = y + rowH * (idx + 1);
        if (yy > pdfHeight - margin) {
          pdf.addPage();
          y = margin;
        }
        pdf.text(r[0], col1X, y + rowH * (idx + 1));
        pdf.text(r[1], col2X, y + rowH * (idx + 1));
      });
      return y + rowH * rows.length + lineHeight;
    };

    // Helper to draw values table
    const renderValuesTable = (table: any, startY: number, title?: string): number => {
      if (!table || !table.headers || !table.rows || table.rows.length === 0) return startY;
      let y = startY;
      if (y + 20 > pdfHeight - margin) {
        pdf.addPage();
        y = margin;
      }
      if (title) {
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(12);
        pdf.text(title, margin, y);
        y += 6;
      }
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      const colCount = table.headers.length;
      const tableWidth = pdfWidth - margin * 2;
      const colWidth = tableWidth / colCount;
      const rowH = 6;
      table.headers.forEach((h: string, i: number) => {
        const x = margin + i * colWidth;
        pdf.text(String(h), x, y);
      });
      y += rowH;
      for (let r = 0; r < table.rows.length; r++) {
        const row = table.rows[r];
        if (y + rowH > pdfHeight - margin) {
          pdf.addPage();
          y = margin;
          table.headers.forEach((h: string, i: number) => {
            const x = margin + i * colWidth;
            pdf.text(String(h), x, y);
          });
          y += rowH;
        }
        row.forEach((cell: string, i: number) => {
          const x = margin + i * colWidth;
          pdf.text(String(cell), x, y);
        });
        y += rowH;
      }
      return y + lineHeight;
    };

    // Special handling for grid mode with exactly 2 charts: place both on a single page
    if (options.layoutMode === 'grid' && chartElementIds.length === 2) {
      const dateStr = new Date().toLocaleString();
      const headerY = margin + 6;
      if (options.titleText) {
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(16);
        pdf.text(String(options.titleText), pdfWidth / 2, headerY, { align: 'center' });
      }
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      let metaY = options.titleText ? headerY + lineHeight : headerY;
      pdf.text(`Date: ${dateStr}`, margin, metaY);
      metaY += lineHeight;

      const halfWidth = (pdfWidth - margin * 3) / 2;
      const yStart = metaY + 2;

      for (let i = 0; i < 2; i++) {
        const chartElement = document.getElementById(chartElementIds[i]);
        if (!chartElement) continue;
        const canvasElement = chartElement.querySelector('canvas') as HTMLCanvasElement | null;
        if (!canvasElement) continue;
        const imgData = canvasElement.toDataURL('image/png');
        const imgHeight = (canvasElement.height * halfWidth) / canvasElement.width;
        const x = margin + i * (halfWidth + margin);
        const t = options.titles?.[i];
        if (t) {
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(11);
          pdf.text(String(t), x + halfWidth / 2, yStart, { align: 'center' });
        }
        pdf.addImage(imgData, 'PNG', x, yStart + 3, halfWidth, imgHeight);
      }

      // Tables on next page
      pdf.addPage();
      for (let i = 0; i < 2; i++) {
        const s = options.summaries?.[i] || null;
        const vt = options.tooltipTablesArr?.[i] || null;
        let y = margin;
        if (options.titles?.[i]) {
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(12);
          pdf.text(String(options.titles[i]), margin, y);
          y += 6;
        }
        if (s) y = renderSummary(s, y, 'Data Summary');
        if (vt) y = renderValuesTable(vt, y, 'Data Values');
        if (i === 0) {
          pdf.addPage();
        }
      }

      pdf.setProperties({
        title: options.titleText || 'Data Visualization Charts',
        subject: 'Multiple Charts Export',
        author: 'BPS Data Visualization',
        creator: 'BPS Chart Dashboard'
      });

      pdf.save(filename);
      console.log(`✅ 2 charts exported on a single page (grid mode): ${filename}`);
      return;
    }

    // Default/list mode: one chart per page, include per-chart title and tables
    for (let i = 0; i < chartElementIds.length; i++) {
      const chartElementId = chartElementIds[i];
      const chartElement = document.getElementById(chartElementId);
      if (!chartElement) {
        console.warn(`⚠️ Chart element with ID '${chartElementId}' not found, skipping...`);
        continue;
      }
      const canvasElement = chartElement.querySelector('canvas') as HTMLCanvasElement | null;
      if (!canvasElement) {
        console.warn(`⚠️ Canvas element not found inside chart with ID '${chartElementId}', skipping...`);
        continue;
      }
      if (i > 0) pdf.addPage();

      const dateStr = new Date().toLocaleString();
      const headerY = margin + 6;
      if (options.titles?.[i]) {
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(16);
        pdf.text(String(options.titles[i]), pdfWidth / 2, headerY, { align: 'center' });
      }
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      let metaY = options.titles?.[i] ? headerY + lineHeight : headerY;
      if (options.sourceNames?.[i]) {
        pdf.text(`Generated From: ${options.sourceNames[i]}`, margin, metaY);
        metaY += lineHeight;
      }
      pdf.text(`Date: ${dateStr}`, margin, metaY);
      metaY += lineHeight;
      if (options.chartTypesArr?.[i]) {
        pdf.text(`Chart Type: ${options.chartTypesArr[i]}`, margin, metaY);
        metaY += lineHeight;
      }

      const imageY = metaY + 2;
      const imgWidth = pdfWidth - margin * 2;
      let imgHeight = (canvasElement.height * imgWidth) / canvasElement.width;
      const imgData = canvasElement.toDataURL('image/png');
      const available = pdfHeight - imageY - lineHeight - margin;
      if (imgHeight > available) {
        const ratio = available / imgHeight;
        const w = imgWidth * ratio;
        const h = imgHeight * ratio;
        pdf.addImage(imgData, 'PNG', margin, imageY, w, h);
        imgHeight = h;
      } else {
        pdf.addImage(imgData, 'PNG', margin, imageY, imgWidth, imgHeight);
      }

      let y = imageY + imgHeight + lineHeight;
      const s = options.summaries?.[i] || null;
      const vt = options.tooltipTablesArr?.[i] || null;
      if (s) y = renderSummary(s, y, 'Data Summary');
      if (vt) y = renderValuesTable(vt, y, 'Data Values');
    }

    // Add metadata
    pdf.setProperties({
      title: options.titleText || 'Data Visualization Charts',
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
      quality = 1.5,
      margin = 5
    } = options;

    const dashboardElement = document.getElementById(dashboardElementId);
    if (!dashboardElement) {
      throw new Error(`Dashboard element with ID '${dashboardElementId}' not found`);
    }

    // Capture the entire dashboard with higher scale for clarity
    const effectiveScale = Math.min(3, (window.devicePixelRatio || 1) * (quality || 1));
    const canvas = await html2canvas(dashboardElement, {
      scale: effectiveScale,
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
