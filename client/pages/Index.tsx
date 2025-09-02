import React, { useState, useCallback, useRef, useEffect } from "react";
import { ArrowLeft, Upload, FileSpreadsheet, Search, X, Download, BarChart3, LineChart, PieChart, ScatterChart, Trash2, Eye, RefreshCw, AlertCircle, CheckCircle } from "lucide-react";
import * as XLSX from 'xlsx';
import Header from "@/components/Header";
import Footer from '@/components/Footer';
import StepIndicator from "@/components/StepIndicator";
import Breadcrumb from "@/components/Breadcrumb";
import CollapsibleSection from "@/components/CollapsibleSection";
import ChartDashboard from "@/components/ChartDashboard";
import DataChart from "@/components/DataChart";
import { exportToExcel } from "@/components/ExportUtils";
import { cn } from "@/lib/utils";

interface ExcelSheet {
  name: string;
  rows: number;
  columns: number;
  data: any[][];
}

interface SelectedTable {
  id: string;
  name: string;
  years: number[];
  previewData: any[][];
}

interface DataSelection {
  tables: SelectedTable[];
  selectedYears: number[];
}

const STEPS = [
  "Upload File",
  "Pilih Tabel", 
  "Filter Tahun",
  "Preview Data",
  "Hasil Tabel",
  "Visualisasi"
];

const CHART_TYPES = [
  { id: "bar", name: "Bar Chart", icon: BarChart3 },
  { id: "line", name: "Line Chart", icon: LineChart },
  { id: "pie", name: "Pie Chart", icon: PieChart },
  { id: "scatter", name: "Scatter Plot", icon: ScatterChart },
];

const SAMPLE_TABLES = [
  {
    id: "tpak-2020-2023",
    name: "Tingkat Partisipasi Angkatan Kerja (TPAK)",
    years: [2020, 2021, 2022, 2023],
    previewData: [
      ["Kabupaten/Kota", "2020", "2021", "2022", "2023"],
      ["SUMATERA UTARA", "65.50", "67.20", "68.22", "71.08"],
      ["NIAS", "62.30", "63.50", "64.00", "69.69"],
      ["MANDAILING NATAL", "69.80", "70.45", "71.15", "63.07"],
      ["TAPANULI SELATAN", "66.20", "67.80", "68.90", "70.15"],
      ["TAPANULI TENGAH", "64.50", "65.90", "66.75", "68.20"]
    ]
  },
  {
    id: "penduduk-kerja-2020-2023", 
    name: "Penduduk Bekerja 15 Tahun ke Atas",
    years: [2020, 2021, 2022, 2023],
    previewData: [
      ["Kabupaten/Kota", "2020", "2021", "2022", "2023"],
      ["SUMATERA UTARA", "7000.5", "7150.2", "7284.1", "7456.2"],
      ["NIAS", "135.2", "138.5", "142.8", "145.6"],
      ["MANDAILING NATAL", "85.3", "87.2", "89.7", "92.1"],
      ["TAPANULI SELATAN", "420.5", "430.1", "438.9", "445.2"],
      ["TAPANULI TENGAH", "380.2", "388.5", "395.3", "402.1"]
    ]
  },
  {
    id: "pengangguran-2020-2023",
    name: "Tingkat Pengangguran Terbuka (TPT)",
    years: [2020, 2021, 2022, 2023], 
    previewData: [
      ["Kabupaten/Kota", "2020", "2021", "2022", "2023"],
      ["SUMATERA UTARA", "6.50", "6.20", "5.83", "5.42"],
      ["NIAS", "3.80", "3.50", "3.21", "3.08"],
      ["MANDAILING NATAL", "5.20", "4.95", "4.67", "4.23"],
      ["TAPANULI SELATAN", "4.80", "4.60", "4.35", "4.10"],
      ["TAPANULI TENGAH", "5.10", "4.85", "4.55", "4.25"]
    ]
  }
];
// ...existing code...

const filterTableDataByYears = (tableData: any[][], selectedYears: number[]) => {
  console.log(`🔍 FILTERING DATA FOR YEARS: ${selectedYears.join(", ")}`);

  if (!tableData || tableData.length === 0 || selectedYears.length === 0) {
    console.log("⚠️ No data or years to filter, returning original data");
    return tableData;
  }

  // Check if we have at least one row
  if (tableData.length < 1) {
    console.log("��️ No rows in table data");
    return tableData;
  }

  // Find headers - might be in first OR second row
  let headerRowIndex = 0;
  let headers = tableData[0];

  // If first row doesn't contain years, check second row
  const firstRowYearCount = headers.filter(h => {
    const num = parseInt(String(h).trim());
    return !isNaN(num) && num >= 2000 && num <= 2030;
  }).length;

  if (firstRowYearCount === 0 && tableData.length > 1) {
    console.log("��� First row has no years, checking second row...");
    headerRowIndex = 1;
    headers = tableData[1];
  }

  console.log(`📋 Using row ${headerRowIndex} as headers: ${headers.join(" | ")}`);

  // Find which columns contain the selected years
  const yearColumnIndices: number[] = [];

  headers.forEach((header, index) => {
    const headerStr = String(header).trim();

    // Always include first column (usually contains identifiers like "Jenis Pengeluaran")
    if (index === 0) {
      yearColumnIndices.push(index);
      console.log(`✅ Including first column: "${headerStr}"`);
      return;
    }

    // Parse header as year
    const headerYear = parseInt(headerStr);

    // Check if this header is one of our selected years
    if (!isNaN(headerYear) && selectedYears.includes(headerYear)) {
      yearColumnIndices.push(index);
      console.log(`✅ Including year column ${index}: "${headerStr}" (Year: ${headerYear})`);
    } else {
      console.log(`❌ Excluding column ${index}: "${headerStr}" (not in selected years: ${selectedYears.join(", ")})`);
    }
  });

  console.log(`📊 FINAL: Selected ${yearColumnIndices.length} columns: ${yearColumnIndices.map(i => headers[i]).join(", ")}`);

  // Must have at least 2 columns (identifier + at least 1 year)
  if (yearColumnIndices.length < 2) {
    console.log(`❌ ERROR: Only found ${yearColumnIndices.length} columns (need at least 2)`);
    console.log(`Available headers: ${headers.join(", ")}`);
    console.log(`Selected years: ${selectedYears.join(", ")}`);
    console.log(`⚠️ RETURNING ORIGINAL DATA - FILTERING FAILED`);
    return tableData;
  }

  // Filter all rows to only include selected year columns
  const filteredData = tableData.map((row, rowIndex) => {
    const filteredRow = yearColumnIndices.map(colIndex => {
      const value = row[colIndex];
      return value !== undefined && value !== null ? value : "";
    });

    if (rowIndex === 0) {
      console.log(`Row ${rowIndex} filtered: [${row.join(", ")}] → [${filteredRow.join(", ")}]`);
    }

    return filteredRow;
  });

  console.log(`✅ SUCCESS: Filtered from ${headers.length} to ${yearColumnIndices.length} columns`);
  console.log(`✅ New headers: ${filteredData[headerRowIndex]?.join(", ") || "N/A"}`);

  return filteredData;
};
// ...existing code...

export default function Index() {
  const [currentStep, setCurrentStep] = useState(1);

  // Track step changes for debugging
  const setCurrentStepWithLogging = (newStep: number) => {
    console.log(`📍 Step change: ${currentStep} → ${newStep}`);
    if (newStep < currentStep && newStep === 1) {
      console.warn('🚨 Unexpected reset to step 1 detected!');
      console.trace(); // Log stack trace to identify the cause
    }
    setCurrentStep(newStep);
  };
  const [isDragOver, setIsDragOver] = useState(false);
  const [uploadedSheets, setUploadedSheets] = useState<ExcelSheet[]>([]);
  const [selectedTables, setSelectedTables] = useState<string[]>([]);
  const [availableYears, setAvailableYears] = useState<number[]>([2022, 2023]);
  const [selectedYears, setSelectedYears] = useState<number[]>([]);
  const [dataSelection, setDataSelection] = useState<DataSelection>({ tables: [], selectedYears: [] });
  const [searchQuery, setSearchQuery] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [activeChart, setActiveChart] = useState("bar");
  const [resultData, setResultData] = useState<any[][]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [successMessage, setSuccessMessage] = useState("");
  const [selectedRowsForDeletion, setSelectedRowsForDeletion] = useState<{[tableId: string]: Set<number>}>({});
  const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  
  // Save state to session storage (with debounce to prevent infinite loops)
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      const state = {
        currentStep,
        selectedTables,
        selectedYears,
        dataSelection,
        uploadedSheets
      };
      try {
        sessionStorage.setItem('bps-app-state', JSON.stringify(state));
        console.log('💾 State saved to session storage:', { step: currentStep, tables: selectedTables.length, years: selectedYears.length });
      } catch (error) {
        console.error('Error saving state to session storage:', error);
      }
    }, 100); // Debounce by 100ms to prevent rapid saves

    return () => clearTimeout(timeoutId);
  }, [currentStep, selectedTables, selectedYears, dataSelection, uploadedSheets]);

  // Load state from session storage (only once on mount)
  useEffect(() => {
    const loadSavedState = () => {
      try {
        const savedState = sessionStorage.getItem('bps-app-state');
        if (savedState) {
          const state = JSON.parse(savedState);
          console.log('📂 Loading saved state:', state);

          // Only restore state if it's valid and not at step 1
          if (state.currentStep && state.currentStep > 1) {
            setCurrentStepWithLogging(state.currentStep);
            setSelectedTables(state.selectedTables || []);
            setSelectedYears(state.selectedYears || []);
            setDataSelection(state.dataSelection || { tables: [], selectedYears: [] });
            setUploadedSheets(state.uploadedSheets || []);
            console.log('✅ State restored successfully');
          } else {
            console.log('⚠️ Saved state is at step 1 or invalid, not restoring');
          }
        } else {
          console.log('ℹ️ No saved state found');
        }
      } catch (error) {
        console.error('❌ Error loading saved state:', error);
        // Clear corrupted session storage
        sessionStorage.removeItem('bps-app-state');
      }
    };

    loadSavedState();
  }, []); // Empty dependency array ensures this only runs once on mount

  const showError = (message: string) => {
    setErrors([message]);
    setTimeout(() => setErrors([]), 5000);
  };

  const showSuccess = (message: string) => {
    setSuccessMessage(message);
    setTimeout(() => setSuccessMessage(""), 3000);
  };


// Get available years from currently selected tables only
const getAvailableYearsFromSelectedTables = () => {
  if (selectedTables.length === 0) {
    return availableYears; // Return all available years if no tables selected
  }

  const yearsFromSelectedTables = new Set<number>();

  selectedTables.forEach(tableId => {
    if (uploadedSheets.length > 0 && tableId.startsWith('sheet-')) {
      // Handle uploaded sheets
      const sheetIndex = parseInt(tableId.replace('sheet-', ''));
      if (sheetIndex >= 0 && sheetIndex < uploadedSheets.length) {
        const sheet = uploadedSheets[sheetIndex];
        if (sheet.data && sheet.data.length > 0) {
          console.log(`🔍 DETECTING YEARS FROM SHEET: ${sheet.name}`);

          // Check ALL rows, not just first 3, to find all year columns
          sheet.data.forEach((row, rowIndex) => {
            if (rowIndex === 0) {
              // Header row - check each column for year
              row.forEach((header, colIndex) => {
                const headerStr = String(header).trim();

                // Method 1: Direct integer parsing for exact year
                const headerYear = parseInt(headerStr);
                if (!isNaN(headerYear) && headerYear >= 1990 && headerYear <= new Date().getFullYear() + 5) {
                  yearsFromSelectedTables.add(headerYear);
                  console.log(`✅ Found year ${headerYear} in header at column ${colIndex}`);
                }

                // Method 2: Look for years in text like "Data 2020", "Tahun 2021", etc.
                const yearPatterns = [
                  /\b(20[0-9]{2})\b/g,           // 2000-2099
                  /\b(201[0-9])\b/g,             // 2010-2019 specifically
                  /\b(202[0-9])\b/g,             // 2020-2029 specifically
                ];

                yearPatterns.forEach(pattern => {
                  const matches = headerStr.match(pattern);
                  if (matches) {
                    matches.forEach(match => {
                      const year = parseInt(match);
                      if (year >= 1990 && year <= new Date().getFullYear() + 5) {
                        yearsFromSelectedTables.add(year);
                        console.log(`✅ Found year ${year} in header pattern "${headerStr}" at column ${colIndex}`);
                      }
                    });
                  }
                });
              });
            } else if (rowIndex < 5) {
              // Check first few data rows for years that might be in cells
              row.forEach((cell, colIndex) => {
                const cellValue = String(cell).trim();

                // Only check if cell looks like a year (4 digits)
                if (/^\d{4}$/.test(cellValue)) {
                  const cellYear = parseInt(cellValue);
                  if (cellYear >= 1990 && cellYear <= new Date().getFullYear() + 5) {
                    yearsFromSelectedTables.add(cellYear);
                    console.log(`✅ Found year ${cellYear} in data cell at row ${rowIndex}, col ${colIndex}`);
                  }
                }

                // Look for year patterns in text
                const yearMatches = cellValue.match(/\b(20[0-9]{2})\b/g);
                if (yearMatches) {
                  yearMatches.forEach(match => {
                    const year = parseInt(match);
                    if (year >= 1990 && year <= new Date().getFullYear() + 5) {
                      yearsFromSelectedTables.add(year);
                      console.log(`✅ Found year ${year} in data text "${cellValue}" at row ${rowIndex}, col ${colIndex}`);
                    }
                  });
                }
              });
            }
          });

          console.log(`📊 Years detected from sheet "${sheet.name}":`, Array.from(yearsFromSelectedTables).sort((a, b) => a - b));
        }
      }
    } else {
      // Handle sample tables
      const sampleTable = SAMPLE_TABLES.find(t => t.id === tableId);
      if (sampleTable) {
        console.log(`📋 Using years from sample table "${sampleTable.name}":`, sampleTable.years);
        sampleTable.years.forEach(year => yearsFromSelectedTables.add(year));
      }
    }
  });

  const result = Array.from(yearsFromSelectedTables).sort((a, b) => a - b);
  console.log(`🎯 FINAL available years from selected tables:`, result);
  return result;
};

  const validateFileUpload = (files: FileList): boolean => {
    setErrors([]);

    if (files.length === 0) {
      showError("Silakan pilih file untuk diupload");
      return false;
    }

    const file = files[0];

    // Check file size (max 10MB)
    if (file.size > 10 * 1024 * 1024) {
      showError("Ukuran file terlalu besar. Maksimal 10MB");
      return false;
    }

    // Check file type
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'application/excel',
      'application/x-excel',
      'application/x-msexcel'
    ];

    const hasValidExtension = file.name.match(/\.(xlsx|xls)$/i);
    const hasValidMimeType = validTypes.includes(file.type);

    if (!hasValidExtension && !hasValidMimeType) {
      showError(`Format file "${file.name}" tidak didukung. Gunakan file Excel (.xlsx atau .xls)`);
      return false;
    }

    return true;
  };

  const validateTableSelection = (): boolean => {
    if (selectedTables.length === 0) {
      showError("Pilih minimal 1 tabel untuk melanjutkan");
      return false;
    }
    return true;
  };

  const validateYearSelection = (): boolean => {
    if (selectedYears.length === 0) {
      showError("Pilih minimal 1 tahun untuk melanjutkan");
      return false;
    }

    // Check if selected years are available in selected tables
    const availableYearsFromTables = getAvailableYearsFromSelectedTables();
    const invalidYears = selectedYears.filter(year => !availableYearsFromTables.includes(year));

    if (invalidYears.length > 0) {
      showError(`Tahun yang dipilih tidak tersedia di tabel: ${invalidYears.join(", ")}`);
      return false;
    }

    return true;
  };

  const handleFileUpload = useCallback(async (files: FileList) => {
    if (!validateFileUpload(files)) return;

    setIsLoading(true);
    setErrors([]);

    try {
      const file = files[0];

      // Read the file as ArrayBuffer
      const arrayBuffer = await file.arrayBuffer();
      console.log(`Processing file: ${file.name}, size: ${file.size} bytes`);

      // Parse with XLSX
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      console.log(`Workbook loaded with ${workbook.SheetNames.length} sheets:`, workbook.SheetNames);

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        showError("File Excel tidak mengandung sheet data yang valid.");
        return;
      }

      // Store source file name
      setUploadedFileName(file.name);

      // Extract sheets data
      const processedSheets: ExcelSheet[] = workbook.SheetNames.map(sheetName => {
        const worksheet = workbook.Sheets[sheetName];

        // Convert sheet to array of arrays
        const sheetData = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          raw: false
        }) as any[][];

        // Filter out empty rows
        const filteredData = sheetData.filter(row =>
          row.some(cell => cell !== null && cell !== undefined && cell !== "")
        );

        // Get dimensions
        const rows = filteredData.length;
        const columns = filteredData.length > 0 ? Math.max(...filteredData.map(row => row.length)) : 0;

        // Ensure all rows have same number of columns
        const normalizedData = filteredData.map(row => {
          const newRow = [...row];
          while (newRow.length < columns) {
            newRow.push("");
          }
          return newRow;
        });

        return {
          name: sheetName,
          rows,
          columns,
          data: normalizedData
        };
      });

      // Filter out empty sheets
      const validSheets = processedSheets.filter(sheet => sheet.rows > 0 && sheet.columns > 0);

      if (validSheets.length === 0) {
        showError("File tidak mengandung data yang valid. Pastikan file Excel berisi data.");
        return;
      }

      // Auto-detect years from all sheets (headers and first few rows)
      const detectedYears = new Set<number>();
      validSheets.forEach(sheet => {
        if (sheet.data.length > 0) {
          // Check first 3 rows to find years (headers + some data rows)
          const rowsToCheck = sheet.data.slice(0, Math.min(3, sheet.data.length));

          rowsToCheck.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
              const cellValue = String(cell).trim();

              // Multiple year detection methods
              const yearPatterns = [
                /\b(20[0-9]{2})\b/g,           // 2000-2099
                /\b(19[0-9]{2})\b/g,           // 1900-1999
                /^(20[0-9]{2})$/,              // Exact match for year
                /^(19[0-9]{2})$/               // Exact match for year
              ];

              yearPatterns.forEach(pattern => {
                const matches = cellValue.match(pattern);
                if (matches) {
                  matches.forEach(match => {
                    const year = parseInt(match);
                    // Valid year range: 1990 to current year + 5
                    if (year >= 1990 && year <= new Date().getFullYear() + 5) {
                      detectedYears.add(year);
                      console.log(`Detected year ${year} in sheet "${sheet.name}" at row ${rowIndex}, col ${colIndex}: "${cellValue}"`);
                    }
                  });
                }
              });

              // Also try parsing as pure number
              const numValue = parseFloat(cellValue);
              if (!isNaN(numValue) && numValue >= 1990 && numValue <= new Date().getFullYear() + 5 && numValue % 1 === 0) {
                detectedYears.add(numValue);
                console.log(`Detected year ${numValue} as number in sheet "${sheet.name}" at row ${rowIndex}, col ${colIndex}`);
              }
            });
          });
        }
      });

      // Update available years if we detected any
      if (detectedYears.size > 0) {
        const sortedYears = Array.from(detectedYears).sort((a, b) => a - b);
        setAvailableYears(sortedYears);
        setSelectedYears([]); // Clear previous selections
        console.log(`Updated available years to:`, sortedYears);
        showSuccess(`Berhasil memproses ${validSheets.length} sheet data. Terdeteksi ${sortedYears.length} tahun: ${sortedYears.join(", ")}`);
      } else {
        // If no years detected, use a reasonable default range
        const currentYear = new Date().getFullYear();
        const defaultYears = Array.from({length: 15}, (_, i) => currentYear - 14 + i); // Last 15 years
        setAvailableYears(defaultYears);
        setSelectedYears([]);
        console.log(`No years detected, using default range:`, defaultYears);
        showSuccess(`Berhasil memproses ${validSheets.length} sheet data dari file ${file.name}. Menggunakan rentang tahun default.`);
      }

      setUploadedSheets(validSheets);
      setCurrentStepWithLogging(2);

    } catch (error) {
      console.error("Error processing file:", error);
      showError("Gagal memproses file Excel. Pastikan file tidak rusak dan berformat .xlsx atau .xls");
    } finally {
      setIsLoading(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileUpload(e.dataTransfer.files);
    }
  }, [handleFileUpload]);

  const handleTableSelect = (tableId: string) => {
    let newSelectedTables;
    if (selectedTables.includes(tableId)) {
      newSelectedTables = selectedTables.filter(id => id !== tableId);
    } else if (selectedTables.length < 2) {
      newSelectedTables = [...selectedTables, tableId];
    } else {
      return; // Don't change if already at maximum
    }

    setSelectedTables(newSelectedTables);

    // Clear selected years when table selection changes
    // so users can see years available from the newly selected tables
    if (newSelectedTables.length !== selectedTables.length) {
      setSelectedYears([]);
      
      // Update available years based on newly selected tables
      if (newSelectedTables.length > 0) {
        const yearsFromSelectedTables = new Set<number>();

        newSelectedTables.forEach(tableId => {
          if (uploadedSheets.length > 0 && tableId.startsWith('sheet-')) {
            // Handle uploaded sheets
            const sheetIndex = parseInt(tableId.replace('sheet-', ''));
            if (sheetIndex >= 0 && sheetIndex < uploadedSheets.length) {
              const sheet = uploadedSheets[sheetIndex];
              if (sheet.data && sheet.data.length > 0) {
                // Check all rows in the sheet for years
                sheet.data.forEach((row, rowIndex) => {
                  row.forEach((cell, colIndex) => {
                    const cellValue = String(cell).trim();

                    // Multiple year detection methods
                    const yearPatterns = [
                      /\b(20[0-9]{2})\b/g,           // 2000-2099
                      /\b(19[0-9]{2})\b/g,           // 1900-1999
                      /^(20[0-9]{2})$/,              // Exact match for year
                      /^(19[0-9]{2})$/               // Exact match for year
                    ];

                    yearPatterns.forEach(pattern => {
                      const matches = cellValue.match(pattern);
                      if (matches) {
                        matches.forEach(match => {
                          const year = parseInt(match);
                          // Valid year range: 1990 to current year + 5
                          if (year >= 1990 && year <= new Date().getFullYear() + 5) {
                            yearsFromSelectedTables.add(year);
                          }
                        });
                      }
                    });

                    // Also try parsing as pure number
                    const numValue = parseFloat(cellValue);
                    if (!isNaN(numValue) && numValue >= 1990 && numValue <= new Date().getFullYear() + 5 && numValue % 1 === 0) {
                      yearsFromSelectedTables.add(numValue);
                    }
                  });
                });
              }
            }
          } else {
            // Handle sample tables
            const sampleTable = SAMPLE_TABLES.find(t => t.id === tableId);
            if (sampleTable) {
              sampleTable.years.forEach(year => yearsFromSelectedTables.add(year));
            }
          }
        });

        const sortedYears = Array.from(yearsFromSelectedTables).sort((a, b) => a - b);
        if (sortedYears.length > 0) {
          setAvailableYears(sortedYears);
          console.log(`Updated available years based on selected tables:`, sortedYears);
        }
      }
    }
  };

  const handleYearToggle = (year: number) => {
    if (selectedYears.includes(year)) {
      setSelectedYears(selectedYears.filter(y => y !== year));
    } else {
      setSelectedYears([...selectedYears, year]);
    }
  };

  const selectAllYears = () => {
    const currentAvailableYears = getAvailableYearsFromSelectedTables();
    if (selectedYears.length === currentAvailableYears.length) {
      setSelectedYears([]);
    } else {
      setSelectedYears([...currentAvailableYears]);
    }
  };

  const proceedToNextStep = () => {
    setErrors([]);

    try {
      if (currentStep === 2) {
        if (!validateTableSelection()) return;
        showSuccess(`${selectedTables.length} tabel terpilih`);
        setCurrentStepWithLogging(3);
      } else if (currentStep === 3) {
        if (!validateYearSelection()) return;

        // Create data selection with better error handling
        const validTables = [];

        console.log("=== CREATING DATA SELECTION ===");
        console.log("Selected tables:", selectedTables);
        console.log("Selected years:", selectedYears);

        for (const id of selectedTables) {
          let table;

          if (uploadedSheets.length > 0 && id.startsWith('sheet-')) {
            // Handle uploaded sheets
            const sheetIndex = parseInt(id.replace('sheet-', ''));

            // Validate index bounds
            if (sheetIndex >= 0 && sheetIndex < uploadedSheets.length) {
              const sheet = uploadedSheets[sheetIndex];
              console.log(`Processing uploaded sheet: ${sheet.name}`);

              // Filter data by selected years
              const filteredData = filterTableDataByYears(sheet.data, selectedYears);

              table = {
                id: id,
                name: sheet.name,
                years: selectedYears, // Use selected years for uploaded data
                previewData: filteredData
              };
              
              console.log(`Created table for uploaded sheet:`, table);
            }
          } else {
            // Handle sample tables
            const sampleTable = SAMPLE_TABLES.find(t => t.id === id);
            if (sampleTable) {
              console.log(`Processing sample table: ${sampleTable.name}`);
              console.log(`Sample table original data:`, sampleTable.previewData);
              
              // For sample tables, also filter by selected years
              const availableYears = sampleTable.years.filter(year => selectedYears.includes(year));
              const filteredData = filterTableDataByYears(sampleTable.previewData, selectedYears);

              table = {
                id: sampleTable.id,
                name: sampleTable.name,
                years: availableYears,
                previewData: filteredData
              };
              
              console.log(`Created table for sample table:`, table);
            }
          }

          if (!table) {
            console.error(`Failed to find table with ID: ${id}`);
            console.error(`Available uploaded sheets:`, uploadedSheets.map((s, i) => ({ index: i, name: s.name })));
            console.error(`Available sample tables:`, SAMPLE_TABLES.map(t => ({ id: t.id, name: t.name })));
            showError(`Tabel dengan ID ${id} tidak ditemukan. Silakan pilih ulang tabel.`);
            return; // Exit the function if table not found
          }

          validTables.push(table);
        }

        if (validTables.length === 0) {
          showError("Tidak ada tabel valid yang ditemukan. Silakan pilih ulang tabel.");
          return;
        }

        const selection: DataSelection = {
          tables: validTables,
          selectedYears
        };

        setDataSelection(selection);
        showSuccess(`Filter tahun berhasil diterapkan: ${selectedYears.length} tahun`);
        setCurrentStepWithLogging(4);
      } else if (currentStep === 4) {
        // Prepare result data (we'll show individual tables, so no need for complex combining)
        if (dataSelection.tables.length === 0) {
          showError("Tidak ada tabel yang dipilih untuk diproses.");
          return;
        }

        // Just set some basic result data for compatibility, actual display uses dataSelection.tables
        setResultData([]);
        showSuccess(`Data siap ditampilkan dari ${dataSelection.tables.length} tabel`);
        setCurrentStepWithLogging(5);
      } else if (currentStep === 5) {
        showSuccess("Tabel hasil siap untuk visualisasi");
        setCurrentStepWithLogging(6);
      }
    } catch (error) {
      console.error("Error in proceedToNextStep:", error);
      showError("Terjadi kesalahan saat memproses data. Silakan coba lagi.");
    }
  };

  const removeTableSelection = (tableId: string) => {
    const newSelectedTables = selectedTables.filter(id => id !== tableId);
    setSelectedTables(newSelectedTables);

    // Clear selected years when table is removed so user can see years from remaining tables
    setSelectedYears([]);

    // If removing a table from data selection, update that too
    if (dataSelection.tables.some(t => t.id === tableId)) {
      setDataSelection({
        tables: dataSelection.tables.filter(t => t.id !== tableId),
        selectedYears: []
      });
    }
  };


  const clearAllSelections = () => {
    if (selectedTables.length > 0 || selectedYears.length > 0) {
      const confirmed = window.confirm('Apakah Anda yakin ingin menghapus semua pilihan? Tindakan ini tidak dapat dibatalkan.');
      if (!confirmed) return;
    }

    setSelectedTables([]);
    setSelectedYears([]);
    setDataSelection({ tables: [], selectedYears: [] });
    showSuccess('Semua pilihan telah dihapus');
  };

  const resetToStart = () => {
    const confirmed = window.confirm('Apakah Anda yakin ingin reset ke langkah awal? Semua data dan pilihan akan hilang.');
    if (!confirmed) return;

    setCurrentStepWithLogging(1);
    setUploadedSheets([]);
    setUploadedFileName(null);
    clearAllSelectionsWithoutConfirm(); // Use version without confirmation to avoid double prompt
    setSelectedRowsForDeletion({});

    // Clear session storage
    sessionStorage.removeItem('bps-app-state');
    showSuccess('Aplikasi telah direset ke langkah awal');
  };

  // Internal version without confirmation for use in resetToStart
  const clearAllSelectionsWithoutConfirm = () => {
    setSelectedTables([]);
    setSelectedYears([]);
    setDataSelection({ tables: [], selectedYears: [] });
  };

  const toggleRowSelection = (tableId: string, rowIndex: number) => {
    setSelectedRowsForDeletion(prev => {
      const newSelection = { ...prev };
      if (!newSelection[tableId]) {
        newSelection[tableId] = new Set();
      }

      if (newSelection[tableId].has(rowIndex)) {
        newSelection[tableId].delete(rowIndex);
      } else {
        newSelection[tableId].add(rowIndex);
      }

      return newSelection;
    });
  };

  const toggleSelectAllRows = (tableId: string, totalRows: number) => {
    setSelectedRowsForDeletion(prev => {
      const newSelection = { ...prev };
      const currentSelection = newSelection[tableId] || new Set();

      if (currentSelection.size === totalRows) {
        // If all rows are selected, deselect all
        newSelection[tableId] = new Set();
      } else {
        // Select all rows (excluding header row)
        newSelection[tableId] = new Set(Array.from({length: totalRows}, (_, i) => i));
      }

      return newSelection;
    });
  };

  const deleteSelectedRows = () => {
    const updatedTables = dataSelection.tables.map(table => {
      const selectedRows = selectedRowsForDeletion[table.id] || new Set();

      if (selectedRows.size === 0) {
        return table; // No rows selected for this table
      }

      // Keep header row (index 0) and filter out selected rows
      const filteredData = table.previewData.filter((row, index) => {
        if (index === 0) return true; // Always keep header
        return !selectedRows.has(index - 1); // Check if data row (index-1) is selected
      });

      console.log(`🗑️ Deleted ${selectedRows.size} rows from table "${table.name}"`);
      console.log(`Remaining rows: ${filteredData.length - 1} (excluding header)`);

      return {
        ...table,
        previewData: filteredData
      };
    });

    setDataSelection({
      ...dataSelection,
      tables: updatedTables
    });

    // Clear selections after deletion
    setSelectedRowsForDeletion({});

    const totalDeleted = Object.values(selectedRowsForDeletion).reduce((sum, set) => sum + set.size, 0);
    showSuccess(`Berhasil menghapus ${totalDeleted} baris data`);
  };

  const [showTutorial, setShowTutorial] = useState(false);
  const loadSampleData = () => {
    console.log('📋 Loading sample data...');

    // Convert SAMPLE_TABLES into uploaded-like sheets to reuse the same flow
    const sheets = SAMPLE_TABLES.map(t => ({
      name: t.name,
      rows: t.previewData.length,
      columns: t.previewData[0]?.length || 0,
      data: t.previewData
    }));

    // Union of years from all samples
    const yearsSet = new Set<number>();
    SAMPLE_TABLES.forEach(t => t.years.forEach(y => yearsSet.add(y)));
    const years = Array.from(yearsSet).sort((a,b) => a - b);

    setUploadedFileName('Sample Dataset');
    setUploadedSheets(sheets);
    setAvailableYears(years);
    setSelectedYears([...years]);
    // Preselect up to two sample sheets for a quick start
    setSelectedTables(['sheet-0', 'sheet-1'].filter(id => (sheets as any[])[parseInt(id.replace('sheet-',''))]));
    setShowTutorial(true);
    setCurrentStepWithLogging(2);
    showSuccess('Sample data berhasil dimuat. Silakan ikuti tutorial singkat di bawah.');
  };

  // Use uploaded sheets if available, otherwise fall back to sample tables
  const availableTables = uploadedSheets.length > 0 ? uploadedSheets : SAMPLE_TABLES;

  // Get years available from selected tables
  const currentAvailableYears = getAvailableYearsFromSelectedTables();

  const filteredTables = availableTables.filter(table =>
    table.name.toLowerCase().includes(searchQuery.toLowerCase())
  );



  const getBreadcrumbItems = () => {
    const items = [
      { label: "Beranda", href: "/" },
      { label: "Visualisasi Data", isActive: true }
    ];

    if (currentStep > 1) {
      items.push({ label: STEPS[currentStep - 1], isActive: true });
    }

    return items;
  };

  return (
    <div className="min-h-screen bg-bps-gray-50">
      <Header />
      <StepIndicator currentStep={currentStep} steps={STEPS} />

      {/* Breadcrumb Navigation */}
      <div className="border-b border-bps-gray-200 bg-white">
        <div className="bps-container py-3">
          <Breadcrumb items={getBreadcrumbItems()} />
        </div>
      </div>

      {/* Notification Messages */}
      {errors.length > 0 && (
        <div className="bps-container pt-4">
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 mb-4">
            <div className="flex items-start">
              <AlertCircle className="w-5 h-5 text-red-500 mt-0.5 mr-3 flex-shrink-0" />
              <div>
                <h4 className="font-medium text-red-800 mb-1">Terjadi Kesalahan</h4>
                <ul className="text-sm text-red-700 space-y-1">
                  {errors.map((error, index) => (
                    <li key={index}>{error}</li>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        </div>
      )}

      {successMessage && (
        <div className="bps-container pt-4">
          <div className="bg-green-50 border border-green-200 rounded-lg p-4 mb-4">
            <div className="flex items-center">
              <CheckCircle className="w-5 h-5 text-green-500 mr-3" />
              <p className="font-medium text-green-800">{successMessage}</p>
            </div>
          </div>
        </div>
      )}

      <main className="bps-container py-8">
        {/* Quick Actions */}
        <div className="mb-6 flex flex-wrap gap-3">
          {currentStep > 1 && (
            <button
              onClick={() => setCurrentStepWithLogging(currentStep - 1)}
              className="bps-btn-outline text-sm"
            >
              <ArrowLeft className="w-4 h-4 mr-2" />
              Kembali
            </button>
          )}
          <button
            onClick={clearAllSelections}
            className="bps-btn-secondary text-sm"
          >
            <RefreshCw className="w-4 h-4 mr-2" />
            Clear All
          </button>
          <button
            onClick={resetToStart}
            className="bps-btn-secondary text-sm"
          >
            Reset
          </button>
          <button
            onClick={loadSampleData}
            className="bps-btn-primary text-sm"
          >
            Load Sample Data
          </button>
        </div>

        {/* Step 1: File Upload */}
        {currentStep === 1 && (
          <div className="bps-card max-w-2xl mx-auto">
            <div className="bps-card-header">
              <h2 className="text-2xl font-bold text-bps-navy">Upload File Excel</h2>
              <p className="text-bps-gray-600 mt-2">
                Upload file .xlsx untuk mulai analisis data statistik
              </p>
            </div>

            <div
              className={cn(
                "upload-zone",
                isDragOver && "dragover"
              )}
              onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
              onDragLeave={() => setIsDragOver(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files && handleFileUpload(e.target.files)}
                className="hidden"
              />
              
              {isLoading ? (
                <div className="flex flex-col items-center">
                  <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-bps-blue mb-4"></div>
                  <p className="text-bps-gray-600">Memproses file...</p>
                </div>
              ) : (
                <div className="flex flex-col items-center">
                  <Upload className="w-12 h-12 text-bps-gray-400 mb-4" />
                  <p className="text-lg font-medium text-bps-gray-700 mb-2">
                    Drag & drop file Excel atau klik untuk browse
                  </p>
                  <p className="text-sm text-bps-gray-500">
                    Mendukung format .xlsx dan .xls (max 10MB)
                  </p>
                </div>
              )}
            </div>

            {uploadedSheets.length > 0 && (
              <div className="mt-6">
                <h3 className="font-semibold text-bps-navy mb-3">Sheets Terdeteksi:</h3>
                <div className="space-y-2">
                  {uploadedSheets.map((sheet, index) => (
                    <div key={index} className="flex items-center justify-between p-3 bg-bps-gray-50 rounded-md">
                      <div className="flex items-center">
                        <FileSpreadsheet className="w-5 h-5 text-bps-blue mr-3" />
                        <div>
                          <p className="font-medium">{sheet.name}</p>
                          <p className="text-sm text-bps-gray-600">
                            {sheet.rows} rows, {sheet.columns} columns
                          </p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Step 2: Table Selection */}
        {currentStep === 2 && (
          <div className="space-y-6">
            {showTutorial && (
              <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-md">
                <div className="font-semibold text-emerald-800 mb-1">Tutorial Cepat (Sample Data)</div>
                <ul className="text-sm text-emerald-700 list-disc list-inside space-y-1">
                  <li>Langkah 1: Pilih maksimal 2 tabel contoh (sudah dipilih otomatis).</li>
                  <li>Langkah 2: Klik “Lanjut ke Filter Tahun”.</li>
                  <li>Langkah 3: Di halaman berikutnya pilih tahun, lalu lanjutkan hingga Preview dan Visualisasi.</li>
                </ul>
                <button onClick={() => setShowTutorial(false)} className="mt-2 text-xs px-2 py-1 bg-emerald-100 rounded hover:bg-emerald-200">Sembunyikan</button>
              </div>
            )}
            <div className="bps-card">
              <div className="bps-card-header">
                <h2 className="text-2xl font-bold text-bps-navy">Pilih Tabel Data</h2>
                <p className="text-bps-gray-600 mt-2">
                  Pilih maksimal 2 tabel untuk analisis (tersisa: {2 - selectedTables.length})
                </p>
              </div>

              <div className="mb-6">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-bps-gray-400 w-4 h-4" />
                  <input
                    type="text"
                    placeholder="Cari tabel..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="bps-input pl-10"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {filteredTables.map((table, index) => {
                  // Handle both uploaded sheets and sample tables
                  const tableId = 'id' in table ? table.id : `sheet-${index}`;
                  const tableName = table.name;
                  const tableYears = 'years' in table ? table.years : [];
                  const tableData = 'data' in table ? table.data : ('previewData' in table ? table.previewData : []);
                  const rows = 'rows' in table ? table.rows : tableData.length;
                  const columns = 'columns' in table ? table.columns : (tableData[0]?.length || 0);

                  return (
                    <div
                      key={tableId}
                      className={cn(
                        "data-card",
                        selectedTables.includes(tableId) && "selected",
                        selectedTables.length >= 2 && !selectedTables.includes(tableId) && "opacity-50 cursor-not-allowed"
                      )}
                      onClick={() => handleTableSelect(tableId)}
                    >
                      <div className="flex items-start justify-between mb-3">
                        <h3 className="font-semibold text-bps-navy text-sm leading-tight">
                          {tableName}
                        </h3>
                        {selectedTables.includes(tableId) && (
                          <div className="w-6 h-6 bg-bps-blue text-white rounded-full flex items-center justify-center flex-shrink-0 ml-2">
                            <span className="text-xs">✓</span>
                          </div>
                        )}
                      </div>

                      <div className="text-sm text-bps-gray-600 mb-3">
                        {tableYears.length > 0 ? (
                          `Tahun: ${tableYears.join(", ")}`
                        ) : (
                          `${rows} baris, ${columns} kolom`
                        )}
                      </div>

                      <div className="text-xs">
                        <div className="font-medium text-bps-gray-700 mb-2">Preview (3 baris pertama):</div>
                        <div className="bg-bps-gray-50 p-2 rounded text-xs overflow-hidden">
                          {tableData.slice(0, 3).map((row, i) => (
                            <div key={i} className="truncate">
                              {Array.isArray(row) ? row.slice(0, 3).join(" | ") : String(row).slice(0, 50)}
                              {Array.isArray(row) && row.length > 3 ? "..." : ""}
                            </div>
                          ))}
                          {tableData.length === 0 && (
                            <div className="text-bps-gray-400 italic">Tidak ada data preview</div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {selectedTables.length > 0 && (
                <div className="mt-8 pt-6 border-t border-bps-gray-200">
                  <div className="flex items-center gap-3">
                    {currentStep > 1 && (
                      <button
                        onClick={() => setCurrentStepWithLogging(currentStep - 1)}
                        className="bps-btn-outline text-sm"
                      >
                        <ArrowLeft className="w-4 h-4 mr-2" />
                        Kembali
                      </button>
                    )}
                    <button
                      onClick={proceedToNextStep}
                      className="bps-btn-primary text-sm"
                    >
                      Lanjut ke Filter Tahun
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Step 3: Year Filter */}
        {currentStep === 3 && (
          <div className="bps-card max-w-4xl mx-auto">
            <div className="bps-card-header">
              <h2 className="text-2xl font-bold text-bps-navy">Filter Tahun</h2>
              <p className="text-bps-gray-600 mt-2">
                Pilih tahun yang ingin dianalisis ({selectedYears.length} tahun terpilih)
              </p>
              {showTutorial && (
                <div className="mt-3 p-3 bg-emerald-50 border border-emerald-200 rounded-md text-sm text-emerald-800">
                  Pilih semua tahun yang ingin Anda tampilkan, kemudian klik “Lanjut ke Preview Data”.
                </div>
              )}
              {currentAvailableYears.length > 0 && (
                <div className="mt-3 p-3 bg-blue-50 border border-blue-200 rounded-md">
                  <p className="text-sm text-blue-800 font-medium mb-1">
                    📊 Tahun Terdeteksi Otomatis dari Data Excel
                  </p>
                  <p className="text-sm text-blue-700">
                    Sistem telah mendeteksi <strong>{currentAvailableYears.length} tahun</strong> dari tabel yang dipilih: 
                    <strong> {currentAvailableYears.join(", ")}</strong>
                  </p>
                  <p className="text-xs text-blue-600 mt-1">
                    Range: {Math.min(...currentAvailableYears)} - {Math.max(...currentAvailableYears)} 
                    ({Math.max(...currentAvailableYears) - Math.min(...currentAvailableYears) + 1} tahun)
                  </p>
                </div>
              )}
            </div>

            <div className="mb-6 flex flex-wrap gap-3">
              <button
                onClick={() => {
                  if (selectedYears.length === currentAvailableYears.length) {
                    setSelectedYears([]);
                  } else {
                    setSelectedYears([...currentAvailableYears]);
                  }
                }}
                className="bps-btn-secondary text-sm"
              >
                {selectedYears.length === currentAvailableYears.length ? "Deselect All" : "Select All Years"}
              </button>
              
                              
                             </div>

            {currentAvailableYears.length > 0 ? (
              <div>
                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-3">
                  {currentAvailableYears.map((year) => (
                    <label
                      key={year}
                      className="flex items-center space-x-2 p-3 bg-white border border-bps-gray-200 rounded-lg hover:bg-bps-gray-50 cursor-pointer transition-colors"
                    >
                      <input
                        type="checkbox"
                        checked={selectedYears.includes(year)}
                        onChange={() => handleYearToggle(year)}
                        className="year-checkbox"
                      />
                      <span className="font-medium text-bps-gray-700 text-sm">{year}</span>
                    </label>
                  ))}
                </div>
                
              </div>
            ) : (
              <div className="text-center py-8">
                <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-6">
                  <div className="text-yellow-800 mb-3">
                    <p className="font-medium mb-2">⚠️ Tidak Ada Tahun Terdeteksi</p>
                    <p className="text-sm">Sistem tidak dapat menemukan tahun dalam data Excel yang dipilih.</p>
                  </div>
                  <div className="text-sm text-yellow-700 space-y-1">
                    <p><strong>Kemungkinan penyebab:</strong></p>
                    <ul className="list-disc list-inside space-y-1 ml-4">
                      <li>Format tahun tidak standar (contoh: "2023", "Tahun 2023", "Data 2023")</li>
                      <li>Tahun berada di luar rentang 1990-{new Date().getFullYear() + 5}</li>
                      <li>Data Excel tidak mengandung informasi tahun</li>
                      <li>Format sel Excel tidak dikenali sebagai angka</li>
                    </ul>
                  </div>
                  <div className="mt-4 text-xs text-yellow-600">
                    <p><strong>Tips:</strong> Pastikan kolom tahun dalam Excel berisi angka tahun (contoh: 2020, 2021, 2022) atau teks yang mengandung tahun.</p>
                  </div>
                </div>
              </div>
            )}

            {selectedYears.length > 0 && (
              <div className="mt-8 pt-6 border-t border-bps-gray-200">
                <div className="mb-4">
                  <div className="text-sm text-bps-gray-600">
                    <strong>Range Tahun:</strong> {Math.min(...selectedYears)} - {Math.max(...selectedYears)}
                  </div>
                  <div className="text-sm text-bps-gray-600">
                    <strong>Jumlah Tahun:</strong> {selectedYears.length} tahun
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  {currentStep > 1 && (
                    <button
                      onClick={() => setCurrentStepWithLogging(currentStep - 1)}
                      className="bps-btn-outline text-sm"
                    >
                      <ArrowLeft className="w-4 h-4 mr-2" />
                      Kembali
                    </button>
                  )}
                  <button
                    onClick={proceedToNextStep}
                    className="bps-btn-primary text-sm"
                  >
                    Lanjut ke Preview Data
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Step 4: Data Preview */}
        {currentStep === 4 && (
          <div className="bps-card">
            <div className="bps-card-header">
              <div className="flex items-start justify-between">
                <div>
                  <h2 className="text-2xl font-bold text-bps-navy">Data Terpilih</h2>
                  <p className="text-bps-gray-600 mt-2">
                    Review data yang akan dianalisis sebelum melanjutkan
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => {
                      console.log("🧹 CLEANING CHART DATA (remove footer/empty rows only)...");

                      const cleanedTables = dataSelection.tables.map(table => {
                        const cleanedData = table.previewData.filter((row, rowIndex) => {
                          // Always keep the first row (headers)
                          if (rowIndex === 0) return true;

                          const firstColRaw = row?.[0];
                          const firstCol = String(firstColRaw ?? '').trim().toLowerCase();

                          // Determine if row is a footer note or generic label (do NOT remove years)
                          const isFooterLabel = ['catatan', 'note', 'notes', 'keterangan', 'remark', 'remarks', 'sumber', 'footer']
                            .some(keyword => firstCol === keyword || firstCol.startsWith(`${keyword}:`));

                          // Remove rows where all data cells (excluding first col) are empty strings
                          const isEmptyRow = row.slice(1).every(cell => String(cell ?? '').trim() === '');

                          if (isFooterLabel || isEmptyRow) {
                            console.log(`🗑️ Removing footer/empty row from ${table.name}:`, row);
                            return false;
                          }

                          return true;
                        });

                        console.log(`✅ Cleaned ${table.name}: ${table.previewData.length} → ${cleanedData.length} rows`);

                        return {
                          ...table,
                          previewData: cleanedData
                        };
                      });

                      setDataSelection({
                        ...dataSelection,
                        tables: cleanedTables
                      });

                      showSuccess("Data chart dibersihkan: baris footer/ kosong dihapus. Baris tahun tidak dihapus.");
                    }}
                    className="bps-btn-primary text-xs"
                    title="Bersihkan data chart dari footer/kosong (tidak menghapus baris tahun)"
                  >
                    🧹 Clean Chart Data
                  </button>
                  <button
                    onClick={() => setCurrentStepWithLogging(3)}
                    className="bps-btn-secondary text-sm"
                  >
                    ← Edit Filter Tahun
                  </button>
                </div>
              </div>
            </div>

       {/* Year Filter Indicator */}
             {dataSelection.selectedYears.length > 0 && (
               <div className="mt-3 p-3 bg-blue-50 border border-blue-200 rounded-md">
                 <p className="text-sm text-blue-800 font-medium">
                   🔍 Filter Tahun: {dataSelection.selectedYears.join(", ")}
                 </p>
                 <p className="text-xs text-blue-600 mt-1">
                   Data ditampilkan hanya untuk tahun yang dipilih (kolom tahun lainnya telah difilter)
                 </p>
                  <div className="mt-2 p-2 bg-white rounded border text-xs">
                    <p className="font-medium text-blue-700">Kolom yang ditampilkan:</p>
                    <p className="text-blue-600">• Kolom 1: Nama Wilayah (selalu ditampilkan)</p>
                    {dataSelection.selectedYears.map(year => (
                      <p key={year} className="text-blue-600">• Kolom {year}: Data tahun {year}</p>
                    ))}
                                         <button
                       onClick={() => {
                         console.log("=== VERIFYING FILTERED DATA ===");
                         console.log("Data Selection:", dataSelection);
                         
                         if (dataSelection.tables.length === 0) {
                           console.log("❌ No tables in data selection");
                           alert("Tidak ada data yang dipilih. Silakan pilih tabel dan tahun terlebih dahulu.");
                           return;
                         }
                         
                         dataSelection.tables.forEach((table, index) => {
                           console.log(`\n--- Table ${index + 1}: ${table.name} ---`);
                           console.log("Selected years:", dataSelection.selectedYears);
                           console.log("Table years:", table.years);
                           console.log("Data structure:");
                           console.log("  Rows:", table.previewData.length);
                           console.log("  Columns:", table.previewData[0]?.length || 0);
                           console.log("  Headers:", table.previewData[0]);
                           
                           // Verify that only selected years are present
                           const headers = table.previewData[0] || [];
                           const yearColumns = headers.slice(1); // Skip first column (identifier)
                           console.log("  Year columns found:", yearColumns);
                           
                           const expectedYears = dataSelection.selectedYears.map(y => String(y));
                           const actualYears = yearColumns.filter(col => expectedYears.includes(String(col)));
                           
                           console.log("  Expected years:", expectedYears);
                           console.log("  Actual years in filtered data:", actualYears);
                           
                           if (actualYears.length === expectedYears.length) {
                             console.log("✅ Filtering verification PASSED");
                           } else {
                             console.log("❌ Filtering verification FAILED");
                             console.log("  Missing years:", expectedYears.filter(y => !actualYears.includes(y)));
                             console.log("  Extra years:", actualYears.filter(y => !expectedYears.includes(y)));
                           }
                           
                           console.log("  Sample data rows:");
                           table.previewData.slice(1, 4).forEach((row, rowIndex) => {
                             console.log(`    Row ${rowIndex + 1}: ${row.join(" | ")}`);
                           });
                         });
                         
                         alert("Data verification completed. Check browser console (F12) for detailed results.");
                       }}
                       className="mt-2 px-2 py-1 bg-blue-100 text-blue-700 text-xs rounded hover:bg-blue-200"
                     >
                       🔍 Verify Filtering
                     </button>
                  </div>
               </div>
             )}

            <div className="space-y-6">
              {dataSelection.tables.map((table) => (
                <div key={table.id} className="bg-bps-gray-50 p-4 rounded-lg">
                  <div className="flex items-start justify-between mb-4">
                    <div>
                      <h3 className="font-semibold text-bps-navy">{table.name}</h3>
                      <p className="text-sm text-bps-gray-600">
                        Tahun terpilih: {dataSelection.selectedYears.join(", ")}
                      </p>
                      <p className="text-xs text-bps-gray-500">
                        {table.previewData.length > 0 && table.previewData[0] ?
                          `${table.previewData.length - 1} baris, ${table.previewData[0].length} kolom tahun` :
                          "Tidak ada data"}
                      </p>
                    </div>
                    <button
                      onClick={() => removeTableSelection(table.id)}
                      className="text-bps-error hover:bg-white p-2 rounded-md transition-colors"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="bps-table text-xs">
                      <tbody>
                        {table.previewData.map((row, i) => (
                          <tr key={i}>
                            {row.map((cell, j) => (
                              <td key={j} className={i === 0 ? "font-semibold bg-bps-gray-100" : ""}>
                                {cell}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}

              {/* Preview Charts */}
              <div className="mt-6 pt-6 border-t border-bps-gray-200">
                <h3 className="font-semibold text-bps-navy mb-4">Preview Visualisasi</h3>
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  {dataSelection.tables.slice(0, 2).map((table) => (
                    <div key={table.id} className="bg-white border border-gray-200 rounded-lg p-4">
                      <h4 className="font-medium text-gray-900 text-sm mb-2">{table.name}</h4>
                      <DataChart
                        type="bar"
                        data={table.previewData}
                        title={`${table.name} (${dataSelection.selectedYears.join(", ")})`}
                        height={250}
                      />
                    </div>
                  ))}
                </div>
                
                <div className="mt-6 flex justify-center">
                  <div className="flex items-center gap-3">
                    {currentStep > 1 && (
                      <button
                        onClick={() => setCurrentStepWithLogging(currentStep - 1)}
                        className="bps-btn-outline text-sm"
                      >
                        <ArrowLeft className="w-4 h-4 mr-2" />
                        Kembali
                      </button>
                    )}
                    <button
                      onClick={proceedToNextStep}
                      className="bps-btn-primary text-sm"
                    >
                      <Eye className="w-4 h-4 mr-2" />
                      Buat Tabel Hasil
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Step 5: Results Table */}
        {currentStep === 5 && (
          <div className="bps-card">
                         <div className="bps-card-header">
               <div className="flex items-center justify-between">
                 <div>
                   <h2 className="text-2xl font-bold text-bps-navy">Hasil Tabel</h2>
                   <p className="text-bps-gray-600 mt-2">
                     Data dari {dataSelection.tables.length} tabel yang dipilih
                   </p>
                   {dataSelection.selectedYears.length > 0 && (
                     <div className="mt-2 p-2 bg-blue-50 border border-blue-200 rounded text-xs">
                       <p className="text-blue-700 font-medium">
                         📊 Data telah difilter untuk tahun: {dataSelection.selectedYears.join(", ")}
                       </p>
                       <p className="text-blue-600">
                         Hanya menampilkan kolom tahun yang dipilih user
                       </p>
                     </div>
                   )}
                 </div>
                <div className="flex items-center space-x-3">
                  <button
                    onClick={deleteSelectedRows}
                    disabled={Object.values(selectedRowsForDeletion).every(set => set.size === 0)}
                    className={`text-sm ${Object.values(selectedRowsForDeletion).some(set => set.size > 0) ? 'bps-btn-primary' : 'bps-btn-secondary opacity-50 cursor-not-allowed'}`}
                    title={Object.values(selectedRowsForDeletion).some(set => set.size > 0) ? 'Hapus baris yang dipilih' : 'Pilih baris terlebih dahulu'}
                  >
                    <Trash2 className="w-4 h-4 mr-2" />
                    Hapus Baris ({Object.values(selectedRowsForDeletion).reduce((sum, set) => sum + set.size, 0)})
                  </button>
                  <button 
                    onClick={() => {
                      try {
                        const exportData = { tables: dataSelection.tables, selectedYears: dataSelection.selectedYears };
                        const filename = exportToExcel(exportData);
                        showSuccess(`Data berhasil diexport ke ${filename}`);
                      } catch (error) {
                        showError('Gagal mengexport data. Silakan coba lagi.');
                      }
                    }}
                    className="bps-btn-primary text-sm"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Export Excel
                  </button>
                </div>
              </div>
            </div>

            {/* Individual Tables Display */}
            <div className="space-y-6">
              {dataSelection.tables.map((table, tableIndex) => (
                <div key={table.id} className="border border-bps-gray-200 rounded-lg p-4">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-lg font-semibold text-bps-navy">{table.name}</h3>
                    <div className="text-sm text-bps-gray-600">
                      {table.previewData.length - 1} baris data
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="bps-table">
                      <thead>
                        {table.previewData[0] && (
                          <tr>
                            <th className="w-12">
                              <input
                                type="checkbox"
                                className="year-checkbox"
                                checked={(selectedRowsForDeletion[table.id]?.size || 0) === (table.previewData.length - 1) && (table.previewData.length > 1)}
                                onChange={() => toggleSelectAllRows(table.id, table.previewData.length - 1)}
                                title="Pilih/batalkan semua baris"
                              />
                            </th>
                            {table.previewData[0].map((header, i) => (
                              <th key={i}>{header}</th>
                            ))}
                          </tr>
                        )}
                      </thead>
                      <tbody>
                        {table.previewData.slice(1).map((row, i) => {
                          const isSelected = selectedRowsForDeletion[table.id]?.has(i) || false;
                          return (
                            <tr key={i} className={isSelected ? 'bg-red-50 border-red-200' : ''}>
                              <td>
                                <input
                                  type="checkbox"
                                  className="year-checkbox"
                                  checked={isSelected}
                                  onChange={() => toggleRowSelection(table.id, i)}
                                  title={`${isSelected ? 'Batalkan pilihan' : 'Pilih'} baris ${i + 1}`}
                                />
                              </td>
                              {row.map((cell, j) => (
                                <td key={j} className={isSelected ? 'text-red-600' : ''}>{cell}</td>
                              ))}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}
            </div>

            {/* Quick Charts Preview */}
            <div className="mt-8 pt-6 border-t border-bps-gray-200">
              <h3 className="font-semibold text-bps-navy mb-4">Quick Charts Preview</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                {dataSelection.tables.map((table) => (
                  <div key={table.id} className="bg-white border border-gray-200 rounded-lg p-4">
                    <h4 className="font-medium text-gray-900 text-sm mb-2">{table.name}</h4>
                    <DataChart
                      type="bar"
                      data={table.previewData}
                      title={`${table.name} (${dataSelection.selectedYears.join(", ")})`}
                      height={200}
                    />
                  </div>
                ))}
              </div>
              
              <div className="flex items-center justify-between">
                <div className="text-sm text-bps-gray-600">
                  Menampilkan {dataSelection.tables.length} tabel individual dengan preview chart
                </div>
                <div className="flex items-center gap-3">
                  {currentStep > 1 && (
                    <button
                      onClick={() => setCurrentStepWithLogging(currentStep - 1)}
                      className="bps-btn-outline text-sm"
                    >
                      <ArrowLeft className="w-4 h-4 mr-2" />
                      Kembali
                    </button>
                  )}
                  <button
                    onClick={proceedToNextStep}
                    className="bps-btn-primary"
                  >
                    Lanjut ke Visualisasi Lengkap
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Step 6: Visualization */}
        {currentStep === 6 && (
          <div className="space-y-6">
            <div className="bps-card">
              <div className="bps-card-header">
                <h2 className="text-2xl font-bold text-bps-navy">Visualisasi Data</h2>
                <p className="text-bps-gray-600 mt-2">
                  Dashboard visualisasi data dari {dataSelection.tables.length} tabel yang dipilih
                </p>
                
                                 {/* Data Overview */}
                 <div className="mt-4 p-4 bg-blue-50 border border-blue-200 rounded-md">
                   <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                     <div className="text-center">
                       <div className="text-lg font-bold text-blue-600">{dataSelection.tables.length}</div>
                       <div className="text-blue-800">Tabel Dipilih</div>
                     </div>
                     <div className="text-center">
                       <div className="text-lg font-bold text-green-600">{dataSelection.selectedYears.length}</div>
                       <div className="text-green-800">Tahun Dipilih</div>
                     </div>
                     <div className="text-center">
                       <div className="text-lg font-bold text-purple-600">{dataSelection.selectedYears.join(", ")}</div>
                       <div className="text-purple-800">Range Tahun</div>
                     </div>
                     <div className="text-center">
                       <div className="text-lg font-bold text-orange-600">
                         {dataSelection.tables.reduce((total, table) => total + (table.previewData.length - 1), 0)}
                       </div>
                       <div className="text-orange-800">Total Data Points</div>
                     </div>
                   </div>
                   {dataSelection.selectedYears.length > 0 && (
                     <div className="mt-3 p-2 bg-white rounded border text-xs">
                       <p className="text-blue-700 font-medium text-center">
                         📊 Visualisasi hanya menampilkan data untuk tahun yang dipilih: {dataSelection.selectedYears.join(", ")}
                       </p>
                     </div>
                   )}
                 </div>
              </div>

              {/* Chart Dashboard */}
              <ChartDashboard
                tables={dataSelection.tables}
                selectedYears={dataSelection.selectedYears}
                sourceFileName={uploadedFileName || undefined}
              />
            </div>
          </div>
        )}
      </main>
      <Footer />
    </div>
  );
}
