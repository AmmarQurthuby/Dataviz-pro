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

    // Tables below image: show Data Values first, then Summary
    let nextSectionY = imageY + imgHeight + lineHeight;

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
      nextSectionY = tableY + lineHeight;
    }

    // Data Summary table
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

    // Abbreviation helpers
    const knownAbbrev: Record<string, string> = {
      'pengeluaran konsumsi rumah tangga': 'PKRT',
      'pengeluaran konsumsi lembaga nonprofit yang melayani rumah tangga': 'PKLNYMRT',
      'produk domestik regional bruto': 'PDRB',
      'tingkat partisipasi angkatan kerja': 'TPAK',
      'tingkat pengangguran terbuka': 'TPT',
    };
    const toAbbrev = (text: string): string => {
      const norm = String(text || '').trim();
      const key = norm.toLowerCase();
      if (knownAbbrev[key]) return knownAbbrev[key];
      const words = norm
        .replace(/\s+/g, ' ')
        .split(' ')
        .filter(Boolean);
      const stop = new Set(['dan','yang','di','ke','dari','untuk','pada','dengan','oleh','kepada','para','antar','yang','yg']);
      const initials = words
        .filter(w => !stop.has(w.toLowerCase()))
        .map(w => w[0] || '')
        .join('')
        .toUpperCase();
      return initials || norm;
    };
    const abbreviateToFit = (text: string, colWidth: number): string => {
      pdf.setFont('helvetica','normal');
      pdf.setFontSize(8);
      const fullWithAbbrev = `${text} (${toAbbrev(text)})`;
      const maxWidth = colWidth - 1;
      if (pdf.getTextWidth(fullWithAbbrev) <= maxWidth) return fullWithAbbrev;
      const ab = toAbbrev(text);
      if (pdf.getTextWidth(ab) <= maxWidth) return ab;
      // Hard trim if even abbreviation is too long
      let s = ab;
      while (s.length > 1 && pdf.getTextWidth(s + '…') > maxWidth) {
        s = s.slice(0, -1);
      }
      return s + (s.endsWith('…') ? '' : '…');
    };

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
          const text = i === 0 ? abbreviateToFit(String(cell), colWidth) : String(cell);
          pdf.text(text, x, y);
        });
        y += rowH;
      }
      return y + lineHeight;
    };

    // Special handling for grid mode with exactly 2 charts: place both on a single page
    if (options.layoutMode === 'grid' && chartElementIds.length === 2) {
      const dateStr = new Date().toLocaleString();
      const headerY = margin + 6;
      // Multi-line comparison title: Perbandingan : (Judul 1) / Vs / (Judul 2)
      const t1 = options.titles?.[0] || '';
      const t2 = options.titles?.[1] || '';
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(14);
      pdf.text(`Perbandingan : ${t1}`, pdfWidth / 2, headerY, { align: 'center' });
      pdf.setFontSize(12);
      pdf.text('Vs', pdfWidth / 2, headerY + lineHeight, { align: 'center' });
      pdf.setFontSize(14);
      pdf.text(`${t2}`, pdfWidth / 2, headerY + lineHeight * 2, { align: 'center' });

      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(10);
      let metaY = headerY + lineHeight * 3;
      pdf.text(`Date: ${dateStr}`, margin, metaY);
      metaY += lineHeight;

      const halfWidth = (pdfWidth - margin * 3) / 2;
      const yStart = metaY + 2;

      const imgHeights: number[] = [];
      for (let i = 0; i < 2; i++) {
        const chartElement = document.getElementById(chartElementIds[i]);
        if (!chartElement) { imgHeights[i] = 0; continue; }
        const canvasElement = chartElement.querySelector('canvas') as HTMLCanvasElement | null;
        if (!canvasElement) { imgHeights[i] = 0; continue; }
        const imgData = canvasElement.toDataURL('image/png');
        const imgHeight = (canvasElement.height * halfWidth) / canvasElement.width;
        const x = margin + i * (halfWidth + margin);
        // No per-chart title above visualization in grid mode
        pdf.addImage(imgData, 'PNG', x, yStart, halfWidth, imgHeight);
        imgHeights[i] = imgHeight;
      }

      // Render per-column summary and values directly below each visualization, same page only
      for (let i = 0; i < 2; i++) {
        const s = options.summaries?.[i] || null;
        const vt = options.tooltipTablesArr?.[i] || null;
        const x = margin + i * (halfWidth + margin);
        let y = yStart + imgHeights[i] + 4;

        // Values table (fit remaining space without adding new page)
        if (vt && vt.headers && vt.rows && vt.rows.length > 0) {
          const rowH = 5;
          const headerH = rowH;
          const footerSpace = 2;
          const available = (pdfHeight - margin) - y - footerSpace;
          const maxRows = Math.max(0, Math.floor((available - headerH) / rowH));

          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(10);
          pdf.text('Data Values', x, y);
          y += 4;

          pdf.setFont('helvetica', 'normal');
          pdf.setFontSize(8);
          const colCount = vt.headers.length;
          const colWidth = halfWidth / colCount;
          // header
          vt.headers.forEach((h, ci) => {
            const headerText = ci === 0 ? abbreviateToFit(String(h), colWidth) : String(h);
            pdf.text(headerText, x + ci * colWidth, y);
          });
          y += rowH;

          const rowsToRender = vt.rows.slice(0, maxRows);
          rowsToRender.forEach((row) => {
            row.forEach((cell, ci) => {
              const text = ci === 0 ? abbreviateToFit(String(cell), colWidth) : String(cell);
              pdf.text(text, x + ci * colWidth, y);
            });
            y += rowH;
          });
          const remaining = vt.rows.length - rowsToRender.length;
          if (remaining > 0) {
            pdf.text(`(+${remaining} more)`, x, y);
          }
          y += 2;
        }

        // Summary block (compact)
        if (s) {
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(10);
          pdf.text('Data Summary', x, y);
          y += 4;
          pdf.setFont('helvetica', 'normal');
          pdf.setFontSize(9);
          const decimals = s.decimals ?? 0;
          const rows: Array<[string, string]> = [
            ['Total Regions', String(s.totalRegions)],
            ['Total Years', String(s.totalYears)],
            ['Data Points', String(s.dataPoints)],
            ['Average', (s.averageValue ?? 0).toFixed(decimals)],
            ['Maximum', (s.maxValue ?? 0).toFixed(decimals)],
            ['Minimum', (s.minValue ?? 0).toFixed(decimals)],
          ];
          const rowH = 5;
          rows.forEach((r) => {
            // two columns within halfWidth
            pdf.text(r[0], x, y);
            pdf.text(r[1], x + halfWidth / 2, y);
            y += rowH;
          });
        }
      }

      pdf.setProperties({
        title: options.titleText || 'Data Visualization Charts',
        subject: 'Multiple Charts Export',
        author: 'BPS Data Visualization',
        creator: 'BPS Chart Dashboard'
      });

      pdf.save(filename);
      console.log(`✅ 2 charts (grid) exported with summaries and values on one page: ${filename}`);
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
      if (vt) y = renderValuesTable(vt, y, 'Data Values');
      if (s) y = renderSummary(s, y, 'Data Summary');
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

    const dateStr = new Date().toLocaleString();
    const headerY = margin + 6;
    const lineHeight = 6;

    let heightRemaining = imgHeight;
    let position = 0;

    // Add image to PDF, splitting across pages if necessary (first page reserves header space if title exists)
    while (heightRemaining > 0) {
      const headerSpace = position === 0 && options.titleText ? (lineHeight * 3) : 0;
      const availablePageHeight = pdfHeight - (margin * 2) - headerSpace;
      const pageHeight = Math.min(heightRemaining, availablePageHeight);

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

        if (position === 0 && options.titleText) {
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(16);
          pdf.text(String(options.titleText), pdfWidth / 2, headerY, { align: 'center' });
          pdf.setFont('helvetica', 'normal');
          pdf.setFontSize(10);
          pdf.text(`Date: ${dateStr}`, margin, headerY + lineHeight);
        }

        const imageY = margin + (position === 0 && options.titleText ? headerSpace : 0);
        pdf.addImage(pageImgData, 'PNG', margin, imageY, imgWidth, pageHeight);
      }

      heightRemaining -= pageHeight;
      position += pageHeight;
    }

    // Optional summary on a new page
    if (options.summary) {
      pdf.addPage();
      let y = margin;
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(12);
      pdf.text('Data Summary', margin, y);
      y += 4;
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
      const rowH = 6;
      rows.forEach((r, idx) => {
        pdf.text(r[0], col1X, y + rowH * (idx + 1));
        pdf.text(r[1], col2X, y + rowH * (idx + 1));
      });
    }

    // Add metadata
    pdf.setProperties({
      title: options.titleText || 'Data Visualization Dashboard',
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
