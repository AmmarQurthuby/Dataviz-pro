import React, { useState } from 'react';
import { Download, Settings, Eye, EyeOff, FileSpreadsheet, FileText, FileImage, Grid, List, ChevronDown } from 'lucide-react';
import DataChart from './DataChart';
import { cn } from '@/lib/utils';
import { exportToExcel, exportToCSV, getDataStatistics } from './ExportUtils';
import { exportChartToPDF, exportMultipleChartsToPDF, exportDashboardToPDF } from './PDFExportUtils';
interface ChartConfig {
  type: 'bar' | 'line' | 'pie' | 'doughnut';
  height: number;
  selectedYearForPie?: number;
}
interface ChartDashboardProps {
  tables: Array<{
    id: string;
    name: string;
    years: number[];
    previewData: any[][];
  }>;
  selectedYears: number[];
  sourceFileName?: string;
}
const ChartDashboard: React.FC<ChartDashboardProps> = ({ tables, selectedYears, sourceFileName }) => {
  // Initialize chart configs for each table
  const [chartConfigs, setChartConfigs] = useState<Record<string, ChartConfig>>(() => {
    const configs: Record<string, ChartConfig> = {};
    tables.forEach(table => {
      configs[table.id] = {
        type: 'bar',
        height: 400,
        selectedYearForPie: selectedYears.length > 0 ? selectedYears[0] : undefined
      };
    });
    return configs;
  });
  const [visibleTables, setVisibleTables] = useState<Set<string>>(new Set(tables.map(t => t.id)));
  const [layoutMode, setLayoutMode] = useState<'grid' | 'list'>('grid');
  const [isExporting, setIsExporting] = useState(false);
  // Update chart configs when tables or selectedYears change
  React.useEffect(() => {
    console.log(`📅 ChartDashboard: selectedYears changed:`, selectedYears);
    setChartConfigs(prevConfigs => {
      const newConfigs = { ...prevConfigs };
      // Add configs for new tables
      tables.forEach(table => {
        if (!newConfigs[table.id]) {
          newConfigs[table.id] = {
            type: 'bar',
            height: 400,
            selectedYearForPie: selectedYears.length > 0 ? selectedYears[0] : undefined
          };
        } else {
          // Update selectedYearForPie if current selection is not available
          const currentConfig = newConfigs[table.id];
          if (selectedYears.length > 0) {
            if (!currentConfig.selectedYearForPie || !selectedYears.includes(currentConfig.selectedYearForPie)) {
              newConfigs[table.id] = {
                ...currentConfig,
                selectedYearForPie: selectedYears[0]
              };
            }
          } else {
            newConfigs[table.id] = {
              ...currentConfig,
              selectedYearForPie: undefined
            };
          }
        }
      });
      // Remove configs for tables that no longer exist
      Object.keys(newConfigs).forEach(tableId => {
        if (!tables.some(table => table.id === tableId)) {
          delete newConfigs[tableId];
        }
      });
      return newConfigs;
    });
  }, [tables, selectedYears]);
  const chartTypes = [
    { id: 'bar', name: 'Bar Chart', description: 'Bagus untuk perbandingan antar kategori' },
    { id: 'line', name: 'Line Chart', description: 'Ideal untuk tren waktu' },
    { id: 'pie', name: 'Pie Chart', description: 'Menampilkan proporsi data per tahun (dapat memilih tahun)' },
    { id: 'doughnut', name: 'Doughnut Chart', description: 'Variasi pie chart yang lebih modern (dapat memilih tahun)' },
  ];
  // Helper functions for individual chart configurations
  const updateChartConfig = (tableId: string, updates: Partial<ChartConfig>) => {
    setChartConfigs(prev => ({
      ...prev,
      [tableId]: {
        ...prev[tableId],
        ...updates
      }
    }));
  };
  const getChartConfig = (tableId: string): ChartConfig => {
    return chartConfigs[tableId] || {
      type: 'bar',
      height: 400,
      selectedYearForPie: selectedYears.length > 0 ? selectedYears[0] : undefined
    };
  };
  const toggleTableVisibility = (tableId: string) => {
    const newVisible = new Set(visibleTables);
    if (newVisible.has(tableId)) {
      newVisible.delete(tableId);
    } else {
      newVisible.add(tableId);
    }
    setVisibleTables(newVisible);
  };
  const exportChart = async (tableId: string) => {
    const table = tables.find(t => t.id === tableId);
    if (table) {
      try {
        setIsExporting(true);
        const filename = `${table.name.replace(/[^a-zA-Z0-9]/g, '_')}_chart_${Date.now()}.pdf`;
        const config = getChartConfig(tableId);
        const summary = getDataSummary(table.previewData);

        // Build tooltip-like values table
        const buildTooltipTable = (): { title: string; headers: string[]; rows: string[][] } | null => {
          if (!table.previewData || table.previewData.length < 2) return null;
          const headers = table.previewData[0];
          const allRows = table.previewData.slice(1);
          const filteredRows = allRows.filter((row) => {
            const firstCol = String(row?.[0] ?? '').trim().toLowerCase();
            if (
              firstCol === '' ||
              firstCol.includes('tahun') ||
              firstCol.includes('year') ||
              /^\d{4}$/.test(firstCol)
            ) {
              return false;
            }
            return true;
          });
          const decimals = summary?.decimals ?? 0;
          const formatVal = (v: any) => {
            const num = parseFloat(String(v ?? 0).replace(/[^\d.-]/g, ''));
            if (isNaN(num)) return '';
            return decimals > 0 ? num.toFixed(decimals) : num.toLocaleString();
          };

          if (config.type === 'pie' || config.type === 'doughnut') {
            if (!config.selectedYearForPie || selectedYears.length === 0) return null;
            const yearIndexInSelection = selectedYears.indexOf(config.selectedYearForPie);
            if (yearIndexInSelection === -1) return null;
            const colIdx = 1 + yearIndexInSelection;
            const rows = filteredRows.map((row, i) => [
              String(row?.[0] ?? `Item ${i + 1}`),
              formatVal(row?.[colIdx])
            ]);
            return {
              title: `Data Values (Tahun ${config.selectedYearForPie})`,
              headers: ['Kategori', String(config.selectedYearForPie)],
              rows
            };
          } else {
            const yearHeaders = selectedYears.map(y => String(y));
            const rows = filteredRows.map((row, i) => {
              const label = String(row?.[0] ?? `Item ${i + 1}`);
              const cells = selectedYears.map((_, idx) => formatVal(row?.[1 + idx]));
              return [label, ...cells];
            });
            return {
              title: 'Data Values',
              headers: ['Kategori', ...yearHeaders],
              rows
            };
          }
        };

        const tooltipTable = buildTooltipTable();

        await exportChartToPDF(`chart-${tableId}`, {
          filename,
          orientation: 'landscape',
          format: 'a4',
          titleText: getChartTitle(table.name, tableId),
          sourceName: sourceFileName || table.name,
          chartType: config.type,
          summary,
          tooltipTable,
        } as any);
        alert(`✅ Chart ${table.name} berhasil diexport ke PDF!`);
      } catch (error) {
        console.error('Export error:', error);
        alert('❌ Gagal mengexport chart. Silakan coba lagi.');
      } finally {
        setIsExporting(false);
      }
    }
  };
  const exportAllVisibleCharts = async () => {
    try {
      setIsExporting(true);
      const selectedTables = tables.filter(t => visibleTables.has(t.id));
      const visibleChartIds = selectedTables.map(t => `chart-${t.id}`);
      const filename = `all_charts_${Date.now()}.pdf`;

      const titles = selectedTables.map(t => getChartTitle(t.name, t.id));
      const sourceNames = selectedTables.map(t => sourceFileName || t.name);
      const chartTypesArr = selectedTables.map(t => getChartConfig(t.id).type);
      const summaries = selectedTables.map(t => getDataSummary(t.previewData));

      const buildTooltipTableFor = (table: typeof selectedTables[number]) => {
        const config = getChartConfig(table.id);
        const summary = getDataSummary(table.previewData);
        const data = table.previewData;
        if (!data || data.length < 2) return null;
        const allRows = data.slice(1);
        const filteredRows = allRows.filter((row) => {
          const firstCol = String(row?.[0] ?? '').trim().toLowerCase();
          if (
            firstCol === '' ||
            firstCol.includes('tahun') ||
            firstCol.includes('year') ||
            /^\d{4}$/.test(firstCol)
          ) {
            return false;
          }
          return true;
        });
        const decimals = summary?.decimals ?? 0;
        const formatVal = (v: any) => {
          const num = parseFloat(String(v ?? 0).replace(/[^\d.-]/g, ''));
          if (isNaN(num)) return '';
          return decimals > 0 ? num.toFixed(decimals) : num.toLocaleString();
        };
        if (config.type === 'pie' || config.type === 'doughnut') {
          if (!config.selectedYearForPie || selectedYears.length === 0) return null;
          const yearIndexInSelection = selectedYears.indexOf(config.selectedYearForPie);
          if (yearIndexInSelection === -1) return null;
          const colIdx = 1 + yearIndexInSelection;
          const rows = filteredRows.map((row, i) => [
            String(row?.[0] ?? `Item ${i + 1}`),
            formatVal(row?.[colIdx])
          ]);
          return {
            title: `Data Values (Tahun ${config.selectedYearForPie})`,
            headers: ['Kategori', String(config.selectedYearForPie)],
            rows
          };
        } else {
          const yearHeaders = selectedYears.map(y => String(y));
          const rows = filteredRows.map((row, i) => {
            const label = String(row?.[0] ?? `Item ${i + 1}`);
            const cells = selectedYears.map((_, idx) => formatVal(row?.[1 + idx]));
            return [label, ...cells];
          });
          return {
            title: 'Data Values',
            headers: ['Kategori', ...yearHeaders],
            rows
          };
        }
      };
      const tooltipTablesArr = selectedTables.map(t => buildTooltipTableFor(t));

      const mainTitle = titles.length === 2 ? `Perbandingan: ${titles[0]} vs ${titles[1]}` : 'Export Charts';

      await exportMultipleChartsToPDF(visibleChartIds, {
        filename,
        orientation: 'portrait',
        format: 'a4',
        titleText: mainTitle,
        layoutMode,
        titles,
        sourceNames,
        chartTypesArr,
        summaries,
        tooltipTablesArr,
      });
      alert(`✅ ${visibleChartIds.length} chart berhasil diexport ke PDF!`);
    } catch (error) {
      console.error('Export error:', error);
      alert('❌ Gagal mengexport charts. Silakan coba lagi.');
    } finally {
      setIsExporting(false);
    }
  };
  const exportDashboard = async () => {
    try {
      setIsExporting(true);
      const filename = `dashboard_${Date.now()}.pdf`;

      // Aggregate summary across visible tables
      const aggregate = (() => {
        let totalRegions = 0;
        let totalYears = selectedYears.length;
        let dataPoints = 0;
        let totalSum = 0;
        let validValues = 0;
        let maxValue = Number.NEGATIVE_INFINITY;
        let minValue = Number.POSITIVE_INFINITY;
        let maxDecimals = 0;
        tables.forEach(table => {
          if (!visibleTables.has(table.id)) return;
          const data = table.previewData;
          if (!data || data.length < 2) return;
          const headers = data[0];
          const rows = data.slice(1).filter((row) => {
            const firstCol = String(row?.[0] ?? '').trim().toLowerCase();
            if (
              firstCol === '' ||
              firstCol.includes('tahun') ||
              firstCol.includes('year') ||
              /^\d{4}$/.test(firstCol)
            ) {
              return false;
            }
            return true;
          });
          totalRegions += rows.length;
          dataPoints += Math.max(0, (headers.length - 1) * rows.length);
          for (let c = 1; c < headers.length; c++) {
            for (let r = 0; r < rows.length; r++) {
              const raw = rows[r]?.[c];
              const s = String(raw ?? '').trim();
              const num = parseFloat(s.replace(/[^0-9\-,.]/g, '').replace(',', '.'));
              if (!isNaN(num)) {
                totalSum += num;
                validValues++;
                maxValue = Math.max(maxValue, num);
                minValue = Math.min(minValue, num);
                const lastDot = s.lastIndexOf('.');
                const lastComma = s.lastIndexOf(',');
                const sepIndex = Math.max(lastDot, lastComma);
                if (sepIndex > -1 && sepIndex < s.length - 1) {
                  maxDecimals = Math.max(maxDecimals, s.length - sepIndex - 1);
                }
              }
            }
          }
        });
        return {
          totalRegions,
          totalYears,
          dataPoints,
          averageValue: validValues > 0 ? totalSum / validValues : 0,
          maxValue: maxValue === Number.NEGATIVE_INFINITY ? 0 : maxValue,
          minValue: minValue === Number.POSITIVE_INFINITY ? 0 : minValue,
          decimals: Math.min(maxDecimals, 3),
        };
      })();

      await exportDashboardToPDF('chart-dashboard', {
        filename,
        orientation: 'portrait',
        format: 'a4',
        titleText: '📊 Dashboard Visualisasi Data',
        summary: aggregate,
        sourceName: sourceFileName,
      });
      alert('✅ Dashboard berhasil diexport ke PDF!');
    } catch (error) {
      console.error('Export error:', error);
      alert('❌ Gagal mengexport dashboard. Silakan coba lagi.');
    } finally {
      setIsExporting(false);
    }
  };
  const handleExportData = (format: 'excel' | 'csv') => {
    try {
      const exportData = { tables, selectedYears };
      let filename = '';
     
      if (format === 'excel') {
        filename = exportToExcel(exportData);
      } else {
        filename = exportToCSV(exportData);
      }
     
      alert(`Data berhasil diexport ke ${filename}`);
    } catch (error) {
      console.error('Export error:', error);
      alert('Gagal mengexport data. Silakan coba lagi.');
    }
  };
  const getChartTitle = (tableName: string, tableId: string) => {
    const config = getChartConfig(tableId);
    if (config.type === 'pie' || config.type === 'doughnut') {
      const yearToShow = config.selectedYearForPie || selectedYears[0];
      return `${tableName} (Tahun ${yearToShow})`;
    }
    return `${tableName} (${selectedYears.join(', ')})`;
  };
  const getDataSummary = (data: any[][]) => {
    if (!data || data.length < 2) return null;
    const headers = data[0];
    // Exclude header-like rows that could contain years instead of data
    const allRows = data.slice(1);
    const filteredRows = allRows.filter((row) => {
      const firstCol = String(row?.[0] ?? '').trim().toLowerCase();
      if (
        firstCol === '' ||
        firstCol.includes('tahun') ||
        firstCol.includes('year') ||
        /^\d{4}$/.test(firstCol)
      ) {
        return false;
      }
      return true;
    });

    const rows = filteredRows;
    const summary = {
      totalRegions: rows.length,
      totalYears: Math.max(0, headers.length - 1),
      dataPoints: Math.max(0, (headers.length - 1) * rows.length),
      averageValue: 0,
      maxValue: 0,
      minValue: Infinity,
      decimals: 0,
    } as { totalRegions: number; totalYears: number; dataPoints: number; averageValue: number; maxValue: number; minValue: number; decimals: number };

    if (rows.length === 0 || headers.length <= 1) {
      summary.minValue = 0;
      return summary;
    }

    let totalSum = 0;
    let validValues = 0;
    let maxDecimals = 0;
    const inferDecimals = (raw: any) => {
      const s = String(raw ?? '').trim();
      const lastDot = s.lastIndexOf('.');
      const lastComma = s.lastIndexOf(',');
      const sepIndex = Math.max(lastDot, lastComma);
      if (sepIndex > -1 && sepIndex < s.length - 1) {
        return s.length - sepIndex - 1;
      }
      return 0;
    };

    // Calculate statistics only from valid data rows
    for (let i = 1; i < headers.length; i++) {
      for (let j = 0; j < rows.length; j++) {
        const raw = rows[j]?.[i];
        const value = parseFloat(String(raw ?? 0).replace(/[^\d.-]/g, ''));
        if (!isNaN(value)) {
          totalSum += value;
          validValues++;
          summary.maxValue = Math.max(summary.maxValue, value);
          summary.minValue = Math.min(summary.minValue, value);
          maxDecimals = Math.max(maxDecimals, inferDecimals(raw));
        }
      }
    }

    summary.averageValue = validValues > 0 ? totalSum / validValues : 0;
    summary.minValue = summary.minValue === Infinity ? 0 : summary.minValue;
    summary.decimals = Math.min(maxDecimals, 3);
    return summary;
  };
  if (tables.length === 0) {
    return (
      <div className="text-center py-12">
        <div className="text-gray-400 text-6xl mb-4">📊</div>
        <h3 className="text-xl font-semibold text-gray-700 mb-2">Belum Ada Data Dipilih</h3>
        <p className="text-gray-500">Pilih tabel dan tahun terlebih dahulu untuk melihat visualisasi</p>
      </div>
    );
  }
  return (
    <div id="chart-dashboard" className="space-y-6 p-4 bg-gray-50 min-h-screen">
      {/* Header Section */}
      <div className="bg-gradient-to-r from-blue-600 to-blue-800 text-white rounded-lg p-6 shadow-lg">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold mb-2">📊 Dashboard Visualisasi Data</h1>
            <p className="text-blue-100">
              {tables.length} tabel • {selectedYears.length} tahun dipilih • {visibleTables.size} chart aktif
            </p>
          </div>
          <div className="text-right">
            <div className="text-2xl font-bold">MULTI-CHART</div>
            <div className="text-blue-200">Mode Individual</div>
          </div>
        </div>
      </div>
      {/* Dashboard Controls */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
               {/* Layout & Visibility Controls */}
               <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
                 <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                   <Eye className="w-5 h-5 mr-2 text-blue-600" />
                   Kontrol Tampilan
                 </h3>
                 
                 <div className="space-y-4">
                   
     
                   <div className="grid grid-cols-2 gap-2">
                     <button
                       onClick={() => setVisibleTables(new Set(tables.map(t => t.id)))}
                       className="flex items-center justify-center px-4 py-2 rounded-lg bg-green-50 text-green-700 hover:bg-green-100 transition-colors text-sm font-medium"
                     >
                       <Eye className="w-4 h-4 mr-1" />
                       Tampilkan Semua
                     </button>
                     <button
                       onClick={() => setVisibleTables(new Set())}
                       className="flex items-center justify-center px-4 py-2 rounded-lg bg-red-50 text-red-700 hover:bg-red-100 transition-colors text-sm font-medium"
                     >
                       <EyeOff className="w-4 h-4 mr-1" />
                       Sembunyikan Semua
                     </button>
                   </div>
                   <div>
                     <label className="block text-sm font-medium text-gray-700 mb-3">Layout Mode</label>
                     <div className="grid grid-cols-2 gap-2">
                       <button
                         onClick={() => setLayoutMode('grid')}
                         className={cn(
                           "flex items-center justify-center px-4 py-3 rounded-xl text-sm font-medium transition-all duration-200",
                           layoutMode === 'grid'
                             ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
                             : "bg-gray-100 text-gray-700 hover:bg-gray-200"
                         )}
                       >
                         <Grid className="w-4 h-4 mr-2" />
                         Grid
                       </button>
                       <button
                         onClick={() => setLayoutMode('list')}
                         className={cn(
                           "flex items-center justify-center px-4 py-3 rounded-xl text-sm font-medium transition-all duration-200",
                           layoutMode === 'list'
                             ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
                             : "bg-gray-100 text-gray-700 hover:bg-gray-200"
                         )}
                       >
                         <List className="w-4 h-4 mr-2" />
                         List
                       </button>
                     </div>
                   </div>
                 </div>
               </div>
     
               {/* Export Controls */}
               <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
                 <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                   <Download className="w-5 h-5 mr-2 text-green-600" />
                   Export Data
                 </h3>
                 
                 <div className="space-y-3">
                   <button
                     onClick={() => handleExportData('excel')}
                     disabled={isExporting}
                     className="w-full flex items-center justify-center px-4 py-3 bg-green-50 text-green-700 rounded-xl hover:bg-green-100 transition-colors text-sm font-medium disabled:opacity-50"
                   >
                     <FileSpreadsheet className="w-4 h-4 mr-2" />
                     Export Excel
                   </button>
                   <button
                     onClick={() => handleExportData('csv')}
                     disabled={isExporting}
                     className="w-full flex items-center justify-center px-4 py-3 bg-blue-50 text-blue-700 rounded-xl hover:bg-blue-100 transition-colors text-sm font-medium disabled:opacity-50"
                   >
                     <FileText className="w-4 h-4 mr-2" />
                     Export CSV
                   </button>
                 </div>
               </div>
     
               {/* PDF Export Controls */}
               <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
                 <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                   <FileImage className="w-5 h-5 mr-2 text-purple-600" />
                   Export PDF
                 </h3>
                 
                 <div className="space-y-3">
                   <button
                     onClick={exportAllVisibleCharts}
                     disabled={isExporting || visibleTables.size === 0}
                     className="w-full flex items-center justify-center px-4 py-3 bg-purple-600 text-white rounded-xl hover:bg-purple-700 transition-colors text-sm font-medium disabled:opacity-50 shadow-lg shadow-purple-200"
                   >
                     <FileImage className="w-4 h-4 mr-2" />
                     {isExporting ? 'Exporting...' : 'Export Charts'}
                   </button>
                   <button
                     onClick={exportDashboard}
                     disabled={isExporting}
                     className="w-full flex items-center justify-center px-4 py-3 bg-orange-600 text-white rounded-xl hover:bg-orange-700 transition-colors text-sm font-medium disabled:opacity-50 shadow-lg shadow-orange-200"
                   >
                     <Download className="w-4 h-4 mr-2" />
                     {isExporting ? 'Exporting...' : 'Export Dashboard'}
                   </button>
                 </div>
               </div>
             </div>

      {/* Charts Grid */}
      <div className={cn(
        "gap-6",
        layoutMode === 'grid'
          ? "grid grid-cols-1 lg:grid-cols-2"
          : "space-y-6"
      )}>
        {tables.map((table) => {
          const isVisible = visibleTables.has(table.id);
          const summary = getDataSummary(table.previewData);
          return (
            <div
              key={table.id}
              id={`chart-${table.id}`}
              className={cn(
                "bg-white rounded-lg border border-gray-200 overflow-hidden transition-all duration-300 shadow-sm hover:shadow-md",
                isVisible ? "opacity-100" : "opacity-50",
                layoutMode === 'list' ? "w-full" : ""
              )}
            >
              {/* Chart Header */}
              <div className="flex items-center justify-between p-4 border-b border-gray-200 bg-gray-50">
                <div className="flex items-center space-x-3">
                  <button
                    onClick={() => toggleTableVisibility(table.id)}
                    className="p-1 rounded hover:bg-gray-100"
                  >
                    {isVisible ? (
                      <Eye className="w-4 h-4 text-gray-600" />
                    ) : (
                      <EyeOff className="w-4 h-4 text-gray-400" />
                    )}
                  </button>
                  <div>
                    <h3 className="font-semibold text-gray-900 text-sm">{table.name}</h3>
                    <p className="text-xs text-gray-500">
                      {summary ? `${summary.totalRegions} wilayah, ${summary.totalYears} tahun` : 'Data tidak tersedia'}
                    </p>
                  </div>
                </div>
               
                <div className="flex items-center space-x-2">
                  {summary && (
                    <div className="text-xs text-gray-500 bg-white px-2 py-1 rounded border">
                      Avg: {summary.averageValue.toFixed(summary.decimals || 0)}
                    </div>
                  )}
                  <button
                    onClick={() => exportChart(table.id)}
                    disabled={isExporting}
                    className="p-2 rounded hover:bg-blue-100 text-blue-600 border border-blue-200 disabled:opacity-50"
                    title="Export Chart ke PDF"
                  >
                    <FileImage className="w-4 h-4" />
                  </button>
                </div>
              </div>
              {/* Chart Content */}
              {isVisible && (() => {
                const config = getChartConfig(table.id);
                console.log(`🎨 Rendering chart for table ${table.id}:`, {
                  chartType: config.type,
                  selectedYearForPie: config.selectedYearForPie,
                  tableYears: table.years,
                  selectedYears,
                  dataRows: table.previewData?.length || 0
                });
                return (
                  <div className="p-6">
                    {/* Individual Chart Controls */}
                    <div className="mb-6 p-4 bg-gray-50 border border-gray-200 rounded-lg">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        {/* Chart Type Selector */}
                        <div>
                          <label className="block text-xs font-medium text-gray-700 mb-2">
                            📊 Jenis Chart
                          </label>
                          <select
                            value={config.type}
                            onChange={(e) => updateChartConfig(table.id, {
                              type: e.target.value as 'bar' | 'line' | 'pie' | 'doughnut'
                            })}
                            className="w-full text-sm border border-gray-300 rounded px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          >
                            {chartTypes.map(type => (
                              <option key={type.id} value={type.id}>
                                {type.name}
                              </option>
                            ))}
                          </select>
                        </div>
                        {/* Height Selector */}
                        <div>
                          <label className="block text-xs font-medium text-gray-700 mb-2">
                            📏 Tinggi Chart
                          </label>
                          <select
                            value={config.height}
                            onChange={(e) => updateChartConfig(table.id, {
                              height: Number(e.target.value)
                            })}
                            className="w-full text-sm border border-gray-300 rounded px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          >
                            <option value={300}>Kecil (300px)</option>
                            <option value={400}>Sedang (400px)</option>
                            <option value={500}>Besar (500px)</option>
                            <option value={600}>Sangat Besar (600px)</option>
                          </select>
                        </div>
                        {/* Year Selector for Pie/Doughnut Charts */}
                        {(config.type === 'pie' || config.type === 'doughnut') && selectedYears.length > 1 && (
                          <div>
                            <label className="block text-xs font-medium text-gray-700 mb-2">
                              🥧 Tahun untuk {config.type === 'pie' ? 'Pie' : 'Doughnut'}
                            </label>
                            <select
                              value={config.selectedYearForPie || selectedYears[0]}
                              onChange={(e) => updateChartConfig(table.id, {
                                selectedYearForPie: Number(e.target.value)
                              })}
                              className="w-full text-sm border border-gray-300 rounded px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                            >
                              {selectedYears.map(year => (
                                <option key={year} value={year}>
                                  Tahun {year}
                                </option>
                              ))}
                            </select>
                          </div>
                        )}
                        {/* Chart Description */}
                        <div className={cn(
                          "col-span-1 md:col-span-3",
                          (config.type === 'pie' || config.type === 'doughnut') && selectedYears.length > 1 ? "md:col-span-1" : ""
                        )}>
                          <div className="text-xs text-gray-600 bg-white p-3 rounded border">
                            <strong>💡 Info:</strong> {chartTypes.find(t => t.id === config.type)?.description}
                          </div>
                        </div>
                      </div>
                    </div>
                    {/* Actual Chart */}
                    <DataChart
                      key={`${table.id}-${config.type}-${config.selectedYearForPie || 'no-year'}`}
                      type={config.type}
                      data={table.previewData}
                      title={getChartTitle(table.name, table.id)}
                      height={config.height}
                      selectedYear={config.selectedYearForPie}
                      selectedYears={selectedYears}  // Tambahan prop ini untuk logika indeks
                    />
                    {/* Data Summary */}
                    {summary && (
                      <div className="mt-6 grid grid-cols-3 gap-4 text-xs">
                        <div className="text-center p-3 bg-gradient-to-b from-red-50 to-red-100 rounded-lg border border-red-200">
                          <div className="font-bold text-red-700 text-lg">{summary.maxValue.toFixed(summary.decimals || 0)}</div>
                          <div className="text-red-600">📈 Maksimum</div>
                        </div>
                        <div className="text-center p-3 bg-gradient-to-b from-green-50 to-green-100 rounded-lg border border-green-200">
                          <div className="font-bold text-green-700 text-lg">{summary.minValue.toFixed(summary.decimals || 0)}</div>
                          <div className="text-green-600">📉 Minimum</div>
                        </div>
                        <div className="text-center p-3 bg-gradient-to-b from-blue-50 to-blue-100 rounded-lg border border-blue-200">
                          <div className="font-bold text-blue-700 text-lg">{summary.dataPoints}</div>
                          <div className="text-blue-600"> Data Points</div>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })()}
            </div>
          );
        })}
      </div>
    </div>
  );
};
export default ChartDashboard;
