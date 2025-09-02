import React from 'react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Title,
  Tooltip,
  Legend,
  Filler,
} from 'chart.js';
import { Bar, Line, Pie, Doughnut } from 'react-chartjs-2';

// Register ChartJS components
ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Title,
  Tooltip,
  Legend,
  Filler
);

interface DataChartProps {
  type: 'bar' | 'line' | 'pie' | 'doughnut';
  data: any[][];
  title?: string;
  height?: number;
  selectedYear?: number; // Untuk pie/doughnut chart memilih tahun yang ditampilkan
  selectedYears?: number[]; // Array tahun yang dipilih untuk menentukan urutan kolom
}

interface ChartData {
  labels: string[];
  datasets: {
    label: string;
    data: number[];
    backgroundColor?: string | string[];
    borderColor?: string | string[];
    borderWidth?: number;
    fill?: boolean;
  }[];
}

const DataChart: React.FC<DataChartProps> = ({ type, data, title, height = 400, selectedYear, selectedYears = [] }) => {
  console.log(`🎯 DataChart component props:`, {
    type,
    title,
    height,
    selectedYear,
    selectedYears,
    dataRows: data?.length || 0,
    dataCols: data?.[0]?.length || 0
  });

  // Force re-render when selectedYear changes for pie charts
  React.useEffect(() => {
    if (type === 'pie' || type === 'doughnut') {
      console.log(`🔄 Pie chart selectedYear changed to:`, selectedYear);
    }
  }, [selectedYear, type]);

  // Helper function to extract pie chart data for specific year
  const extractPieChartData = (headers: any[], dataRows: any[][], yearColumnIndex: number): { labels: string[], values: number[] } => {
    const year = headers[yearColumnIndex] || selectedYears[yearColumnIndex - 1] || 'Unknown';
    console.log(`🥧 Extracting pie chart data for year column ${yearColumnIndex} (${year})`);
    const labels: string[] = [];
    const values: number[] = [];
    dataRows.forEach((row, rowIndex) => {
      const label = String(row[0] || `Item ${rowIndex + 1}`).trim();
      const rawValue = row[yearColumnIndex];
      let value = 0;
      if (rawValue !== null && rawValue !== undefined && rawValue !== '') {
        const cleanValue = String(rawValue).replace(/[^\d.-]/g, '');
        const parsedValue = parseFloat(cleanValue);
        if (!isNaN(parsedValue)) {
          value = parsedValue;
        }
      }
      if (label && !label.toLowerCase().includes('total') && value > 0) {
        labels.push(label);
        values.push(value);
        console.log(`📊 Added pie slice: "${label}" = ${value}`);
      } else {
        console.log(`⚠️ Skipped row ${rowIndex}: label="${label}", value=${value}`);
      }
    });
    console.log(`🥧 Final pie chart data: ${labels.length} slices, total value: ${values.reduce((sum, v) => sum + v, 0)}`);
    return { labels, values };
  };

  // Process data for chart
  const processDataForChart = React.useMemo(() => {
    const timestamp = new Date().toLocaleTimeString();
    console.log(`🎨 DataChart processing data at ${timestamp}:`, {
      type,
      selectedYear,
      selectedYears,
      dataLength: data?.length || 0
    });

    if (!data || data.length < 2) {
      console.log(`❌ DataChart: Insufficient data (${data?.length || 0} rows)`);
      return {
        labels: [],
        datasets: []
      };
    }

    // Identify header row and data rows
    let headers = data[0];
    let dataRows = data.slice(1);

    // Check for misplaced headers
    const firstRowHasOnlyYears = data[0].slice(1).every(cell => {
      const str = String(cell || '').trim();
      return str === '' || /^\d{4}$/.test(str) || /^20\d{2}$/.test(str);
    });

    if (data.length > 1) {
      const secondRowHasYears = data[1].slice(1).every(cell => {
        const str = String(cell || '').trim();
        return str === '' || /^\d{4}$/.test(str) || /^20\d{2}$/.test(str);
      });
      if (firstRowHasOnlyYears && !secondRowHasYears) {
        console.log(`📋 Using row 0 as headers (contains only years)`);
        headers = data[0];
        dataRows = data.slice(1);
      } else if (secondRowHasYears && !firstRowHasOnlyYears) {
        console.log(`📋 Using row 1 as headers (row 0 seems to be data)`);
        headers = data[1];
        dataRows = [data[0], ...data.slice(2)];
      }
    }

    console.log(`📋 Final headers:`, headers);
    console.log(`📊 Final data rows (${dataRows.length}):`, dataRows);

    // Extract labels (first column - usually region names)
    const filteredDataRows = dataRows.filter(row => {
      const firstCol = String(row[0] || '').trim();
      if (firstCol.toLowerCase().includes('tahun') ||
          firstCol.toLowerCase().includes('year') ||
          /^\d{4}$/.test(firstCol) ||
          firstCol === '') {
        console.log(`⚠️ Skipping potential header row: [${row.join(', ')}]`);
        return false;
      }
      return true;
    });

    const labels = filteredDataRows.map((row, index) => {
      const label = String(row[0] || `Row ${index + 1}`);
      return label.trim() || `Item ${index + 1}`;
    });

    console.log(`🏷️ DataChart labels (${labels.length}):`, labels);

    // Extract datasets (year columns)
    const datasets = [];
    const colors = [
      'rgba(30, 58, 138, 0.8)', // BPS Navy
      'rgba(59, 130, 246, 0.8)', // BPS Blue
      'rgba(16, 185, 129, 0.8)', // Green
      'rgba(245, 158, 11, 0.8)', // Orange
      'rgba(239, 68, 68, 0.8)', // Red
      'rgba(139, 92, 246, 0.8)', // Purple
      'rgba(236, 72, 153, 0.8)', // Pink
      'rgba(34, 197, 94, 0.8)', // Light Green
    ];

    if (type === 'pie' || type === 'doughnut') {
      // PIE CHART LOGIC: Use index in selectedYears to find column
      console.log(`🥧 Processing pie chart data for selectedYear: ${selectedYear}`);
      let targetColumnIndex = -1;
      let targetYear = selectedYear;

      if (selectedYears.length > 0 && selectedYear !== undefined) {
        const index = selectedYears.indexOf(selectedYear);
        if (index !== -1) {
          targetColumnIndex = 1 + index;
          console.log(`✅ Found column index ${targetColumnIndex} for year ${selectedYear} (position in selectedYears: ${index})`);
        }
      }

      if (targetColumnIndex === -1) {
        console.error(`❌ Selected year ${selectedYear} not found in selectedYears: ${selectedYears}`);
        return {
          labels: [],
          datasets: []
        };
      }

      const pieData = extractPieChartData(headers, filteredDataRows, targetColumnIndex);
      if (pieData.labels.length > 0 && pieData.values.length > 0) {
        const pieColors = pieData.values.map((_, index) => colors[index % colors.length]);
        datasets.push({
          label: `Data Tahun ${targetYear}`,
          data: pieData.values,
          backgroundColor: pieColors,
          borderColor: pieColors.map(color => color.replace('0.8', '1')),
          borderWidth: 2,
        });
        labels.length = 0;
        labels.push(...pieData.labels);
        console.log(`✅ Pie chart dataset created with ${pieData.labels.length} slices for year ${targetYear}`);
      } else {
        console.error(`❌ No valid pie chart data found for year ${targetYear}`);
      }
    } else {
      // BAR/LINE CHART LOGIC: Process all year columns
      for (let i = 1; i < headers.length; i++) {
        const rawHeader = headers[i];
        // Use selectedYears if available, otherwise extract year from header
        let year: string;
        if (selectedYears.length >= i) {
          year = String(selectedYears[i - 1]);
          console.log(`📅 Using selectedYears[${i - 1}] = ${year} for column ${i}`);
        } else {
          year = String(rawHeader || '').trim();
          const yearMatch = year.match(/\b(20\d{2})\b/);
          if (yearMatch) {
            year = yearMatch[1];
          } else if (year.match(/^\d{4}$/)) {
            year = year;
          } else {
            year = `Tahun ${i}`;
          }
          console.log(`📅 Fallback: Extracted year "${year}" from header "${rawHeader}" for column ${i}`);
        }

        const values = filteredDataRows.map((row, rowIndex) => {
          const rawValue = row[i];
          let value = parseFloat(String(rawValue || 0).replace(/[^\d.-]/g, ''));
          if (isNaN(value)) {
            console.log(`⚠️ DataChart: Invalid value at row ${rowIndex}, col ${i}: "${rawValue}" → 0`);
            value = 0;
          }
          return value;
        });

        console.log(`💹 DataChart values for "${year}" (${values.length} values):`, values);
        const colorIndex = (i - 1) % colors.length;
        datasets.push({
          label: year, // Langsung gunakan tahun sebagai label
          data: values,
          backgroundColor: type === 'bar' ? colors[colorIndex] : 'transparent',
          borderColor: colors[colorIndex],
          borderWidth: 2,
          fill: type === 'line' ? false : undefined,
        });
      }
    }

    console.log(`✅ DataChart final datasets:`, datasets.map(d => ({ label: d.label, dataPoints: d.data.length })));
    return {
      labels,
      datasets
    };
  }, [data, type, selectedYear, selectedYears]);

  // Calculate Y-axis bounds to ensure zero line is visible
  const calculateYAxisBounds = () => {
    if (type === 'pie' || type === 'doughnut') return {};
    const allValues = processDataForChart.datasets.flatMap(dataset => dataset.data);
    if (allValues.length === 0) return { includeZero: true };
    const minValue = Math.min(...allValues);
    const maxValue = Math.max(...allValues);
    console.log(`📊 DataChart Y-axis bounds: min=${minValue}, max=${maxValue}`);
    const range = Math.abs(maxValue - minValue);
    const padding = range > 0 ? range * 0.15 : 1;
    let suggestedMin = minValue;
    let suggestedMax = maxValue;
    if (minValue < 0) {
      suggestedMin = minValue - padding;
    }
    if (maxValue > 0) {
      suggestedMax = maxValue + padding;
    }
    if (minValue > 0) {
      suggestedMin = 0;
    }
    if (maxValue < 0) {
      suggestedMax = 0;
    }
    console.log(`📊 Final Y-axis range: ${suggestedMin} to ${suggestedMax} (includes zero)`);
    return {
      min: suggestedMin,
      max: suggestedMax,
      includeZero: true
    };
  };

  const yAxisBounds = calculateYAxisBounds();

  // Infer maximum number of decimal places from original data values
  const maxDecimals = React.useMemo(() => {
    if (!data || data.length < 2) return 0;
    let maxD = 0;
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!Array.isArray(row)) continue;
      for (let c = 1; c < row.length; c++) {
        const raw = row[c];
        if (raw === null || raw === undefined || raw === '') continue;
        const s = String(raw).trim();
        // Determine decimals from last occurrence of ',' or '.'
        const lastDot = s.lastIndexOf('.');
        const lastComma = s.lastIndexOf(',');
        const sepIndex = Math.max(lastDot, lastComma);
        let dec = 0;
        if (sepIndex > -1 && sepIndex < s.length - 1) {
          dec = s.length - sepIndex - 1;
        }
        // Count only if value is numeric after normalization
        const numeric = parseFloat(s.replace(/[^0-9\-,.]/g, '').replace(',', '.'));
        if (!isNaN(numeric)) {
          maxD = Math.max(maxD, dec);
        }
      }
    }
    // Limit to 3 decimals for readability
    return Math.min(maxD, 3);
  }, [data]);

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'top' as const,
        labels: {
          font: {
            family: 'Inter',
            size: 12,
          },
          color: '#374151',
        },
      },
      title: {
        display: !!title,
        text: title,
        font: {
          family: 'Inter',
          size: 16,
          weight: 'bold',
        },
        color: '#1e3a8a',
      },
      tooltip: {
        backgroundColor: 'rgba(255, 255, 255, 0.95)',
        titleColor: '#1e3a8a',
        bodyColor: '#374151',
        borderColor: '#e5e7eb',
        borderWidth: 1,
        cornerRadius: 8,
        displayColors: true,
        callbacks: {
          label: function(context: any) {
            const label = context.dataset.label || '';
            const value = context.parsed.y || context.parsed;
            if (typeof value === 'number') {
              return `${label}: ${maxDecimals > 0 ? value.toFixed(maxDecimals) : value.toLocaleString()}`;
            }
            return `${label}: ${value}`;
          }
        }
      },
    },
    scales: type === 'bar' || type === 'line' ? {
      x: {
        grid: {
          color: '#e5e7eb',
          display: true,
        },
        ticks: {
          font: {
            family: 'Inter',
            size: 11,
          },
          color: '#6b7280',
        },
      },
      y: {
        ...yAxisBounds,
        position: 'left' as const,
        grid: {
          color: function(context: any) {
            return context.tick.value === 0 ? '#374151' : '#e5e7eb';
          },
          lineWidth: function(context: any) {
            return context.tick.value === 0 ? 3 : 1;
          },
          display: true,
        },
        ticks: {
          font: {
            family: 'Inter',
            size: 11,
          },
          color: '#6b7280',
          callback: (value: any) => {
            if (typeof value === 'number') {
              return maxDecimals > 0 ? value.toFixed(maxDecimals) : value.toLocaleString();
            }
            return value as any;
          }
        },
        border: {
          color: '#374151',
          width: 1,
        },
      },
    } : undefined,
  };

  const renderChart = () => {
    console.log(`🎨 Rendering ${type} chart with data:`, {
      labels: processDataForChart.labels,
      datasets: processDataForChart.datasets.map(d => ({
        label: d.label,
        dataLength: d.data.length,
        data: d.data
      }))
    });

    const chartKey = `${type}-${selectedYear}-${JSON.stringify(processDataForChart.labels)}`;

    if ((type === 'pie' || type === 'doughnut')) {
      if (processDataForChart.datasets.length === 0) {
        console.error(`❌ No datasets available for ${type} chart. SelectedYear: ${selectedYear}`);
        return (
          <div className="flex items-center justify-center h-64 bg-yellow-50 rounded-lg border border-yellow-200">
            <div className="text-center max-w-md">
              <div className="text-yellow-400 text-4xl mb-3">📅</div>
              <p className="text-yellow-800 font-medium mb-2">Data untuk tahun tidak ditemukan</p>
              <p className="text-sm text-yellow-700 mb-3">
                {selectedYear ?
                  `Tahun ${selectedYear} tidak tersedia dalam data ini.` :
                  'Tidak ada tahun yang dipilih untuk pie chart.'
                }
              </p>
              <div className="text-xs text-yellow-600 bg-yellow-100 p-2 rounded">
                💡 <strong>Tips:</strong> Pilih tahun yang tersedia dalam dropdown atau gunakan chart type lain seperti Bar Chart untuk melihat semua tahun.
              </div>
            </div>
          </div>
        );
      }

      if (processDataForChart.labels.length === 0) {
        console.error(`❌ No labels available for ${type} chart`);
        return (
          <div className="flex items-center justify-center h-64 bg-orange-50 rounded-lg border border-orange-200">
            <div className="text-center max-w-md">
              <div className="text-orange-400 text-4xl mb-3">📋</div>
              <p className="text-orange-800 font-medium mb-2">Tidak ada kategori data</p>
              <p className="text-sm text-orange-700 mb-3">
                Data untuk tahun {selectedYear || 'yang dipilih'} tidak memiliki kategori yang valid atau semua nilai adalah nol.
              </p>
              <div className="text-xs text-orange-600 bg-orange-100 p-2 rounded">
                💡 <strong>Solusi:</strong> Pastikan data memiliki kolom kategori (wilayah/sektor) dengan nilai positif.
              </div>
            </div>
          </div>
        );
      }

      const firstDataset = processDataForChart.datasets[0];
      if (firstDataset) {
        const totalValue = firstDataset.data.reduce((sum, value) => sum + (value || 0), 0);
        const validValues = firstDataset.data.filter(value => value > 0).length;
        if (totalValue === 0 || validValues === 0) {
          console.error(`❌ All data values are zero or invalid for ${type} chart`);
          return (
            <div className="flex items-center justify-center h-64 bg-blue-50 rounded-lg border border-blue-200">
              <div className="text-center max-w-md">
                <div className="text-blue-400 text-4xl mb-3">🔢</div>
                <p className="text-blue-800 font-medium mb-2">Semua data bernilai nol</p>
                <p className="text-sm text-blue-700 mb-3">
                  Data untuk tahun {selectedYear || 'yang dipilih'} tidak memiliki nilai positif untuk ditampilkan dalam pie chart.
                </p>
                <div className="text-xs text-blue-600 bg-blue-100 p-2 rounded">
                  💡 <strong>Tips:</strong> Pie chart hanya menampilkan nilai positif. Coba gunakan Bar Chart untuk melihat data termasuk nilai nol.
                </div>
              </div>
            </div>
          );
        }
        console.log(`✅ Pie chart validation passed: ${validValues} valid values out of ${firstDataset.data.length}, total: ${totalValue.toFixed(2)}`);
      }
      console.log(`✅ Pie chart data validation passed. Labels: ${processDataForChart.labels.length}, Datasets: ${processDataForChart.datasets.length}`);
    }

    switch (type) {
      case 'bar':
        return <Bar key={chartKey} data={processDataForChart} options={options} />;
      case 'line':
        return <Line key={chartKey} data={processDataForChart} options={options} />;
      case 'pie':
        return <Pie key={chartKey} data={processDataForChart} options={options} />;
      case 'doughnut':
        return <Doughnut key={chartKey} data={processDataForChart} options={options} />;
      default:
        return <Bar key={chartKey} data={processDataForChart} options={options} />;
    }
  };

  if (!data || data.length < 2) {
    return (
      <div className="flex items-center justify-center h-64 bg-gray-50 rounded-lg">
        <div className="text-center">
          <div className="text-gray-400 text-4xl mb-2">📊</div>
          <p className="text-gray-500">Tidak ada data untuk ditampilkan</p>
          <p className="text-sm text-gray-400">Pastikan data telah dipilih dan diproses</p>
        </div>
      </div>
    );
  }

  return (
    <div style={{ height: `${height}px` }} className="w-full">
      {renderChart()}
    </div>
  );
};

export default DataChart;
