import * as XLSX from 'xlsx';

export interface ExportData {
  tables: Array<{
    id: string;
    name: string;
    years: number[];
    previewData: any[][];
  }>;
  selectedYears: number[];
}

export const exportToExcel = (data: ExportData) => {
  const workbook = XLSX.utils.book_new();
  
  // Create summary sheet
  const summaryData = [
    ['Ringkasan Data Visualisasi'],
    [''],
    ['Total Tabel:', data.tables.length],
    ['Tahun Dipilih:', data.selectedYears.join(', ')],
    ['Total Data Points:', data.tables.reduce((total, table) => total + (table.previewData.length - 1), 0)],
    [''],
    ['Daftar Tabel:'],
    ['ID', 'Nama Tabel', 'Jumlah Tahun', 'Jumlah Data'],
    ...data.tables.map(table => [
      table.id,
      table.name,
      table.years.length,
      table.previewData.length - 1
    ])
  ];
  
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Ringkasan');
  
  // Create individual table sheets
  data.tables.forEach((table, index) => {
    const sheetName = `Tabel_${index + 1}`;
    const sheetData = [
      [`${table.name} (${data.selectedYears.join(', ')})`],
      [''],
      ...table.previewData
    ];
    
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });
  
  // Generate filename
  const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
  const filename = `DataViz_Export_${timestamp}.xlsx`;
  
  // Save file
  XLSX.writeFile(workbook, filename);
  
  return filename;
};

export const exportToCSV = (data: ExportData) => {
  const csvContent = [];
  
  // Add summary
  csvContent.push('Ringkasan Data Visualisasi');
  csvContent.push('');
  csvContent.push(`Total Tabel,${data.tables.length}`);
  csvContent.push(`Tahun Dipilih,${data.selectedYears.join(', ')}`);
  csvContent.push(`Total Data Points,${data.tables.reduce((total, table) => total + (table.previewData.length - 1), 0)}`);
  csvContent.push('');
  
  // Add each table
  data.tables.forEach((table, index) => {
    csvContent.push(`Tabel ${index + 1}: ${table.name} (${data.selectedYears.join(', ')})`);
    csvContent.push('');
    
    // Add headers
    csvContent.push(table.previewData[0].join(','));
    
    // Add data rows
    table.previewData.slice(1).forEach(row => {
      csvContent.push(row.join(','));
    });
    
    csvContent.push('');
  });
  
  const csvString = csvContent.join('\n');
  const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
  
  // Generate filename
  const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
  const filename = `DataViz_Export_${timestamp}.csv`;
  
  // Download file
  const link = document.createElement('a');
  if (link.download !== undefined) {
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
  
  return filename;
};

export const exportChartAsImage = (chartId: string, filename?: string) => {
  const canvas = document.querySelector(`#${chartId} canvas`) as HTMLCanvasElement;
  if (!canvas) {
    console.error('Chart canvas not found');
    return null;
  }
  
  const link = document.createElement('a');
  link.download = filename || `chart_${Date.now()}.png`;
  link.href = canvas.toDataURL('image/png');
  link.click();
  
  return link.download;
};

export const getDataStatistics = (data: ExportData) => {
  const stats = {
    totalTables: data.tables.length,
    totalYears: data.selectedYears.length,
    totalDataPoints: 0,
    averageDataPerTable: 0,
    maxDataPoints: 0,
    minDataPoints: Infinity,
    tableNames: [] as string[],
    yearRange: `${Math.min(...data.selectedYears)} - ${Math.max(...data.selectedYears)}`,
  };
  
  data.tables.forEach(table => {
    const dataPoints = table.previewData.length - 1;
    stats.totalDataPoints += dataPoints;
    stats.maxDataPoints = Math.max(stats.maxDataPoints, dataPoints);
    stats.minDataPoints = Math.min(stats.minDataPoints, dataPoints);
    stats.tableNames.push(table.name);
  });
  
  stats.averageDataPerTable = stats.totalDataPoints / stats.totalTables;
  stats.minDataPoints = stats.minDataPoints === Infinity ? 0 : stats.minDataPoints;
  
  return stats;
};
