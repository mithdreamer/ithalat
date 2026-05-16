/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Plus, 
  ArrowDownCircle, 
  ArrowUpCircle, 
  Package, 
  History, 
  TrendingUp, 
  AlertCircle,
  CheckCircle2,
  Calendar,
  Layers,
  ArrowRight,
  Upload,
  FileSpreadsheet,
  LayoutGrid,
  Link2,
  Box,
  AlertTriangle,
  RotateCcw,
  Search,
  ChevronDown,
  Info,
  FileText,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Import, Export, StockLink, Status } from './types';
import { matchExportWithImports } from './lib/dir-logic';

export default function App() {
  const [imports, setImports] = useState<Import[]>([]);
  const [exports, setExports] = useState<Export[]>([]);
  const [links, setLinks] = useState<StockLink[]>([]);
  const [activeMenu, setActiveMenu] = useState<'overview' | 'links' | 'stock' | 'errors'>('stock');
  const [searchTerm, setSearchTerm] = useState('');
  const [importFileInfo, setImportFileInfo] = useState<{ name: string; rows: number; total: number } | null>(null);
  const [exportFileInfo, setExportFileInfo] = useState<{ name: string; rows: number; total: number } | null>(null);

  const [pendingExportData, setPendingExportData] = useState<any[] | null>(null);
  const [pendingExportFileName, setPendingExportFileName] = useState<string>('');

  const [bulkResults, setBulkResults] = useState<{
    success: boolean;
    message: string;
    items: {
      id?: string;
      declarationNumber: string;
      date: string;
      quantity: number;
      consumptionRate: number;
      itemCode: string;
      status: 'OK' | 'FAIL';
      error?: string;
      linkedImports?: string[];
    }[];
  } | null>(null);

  const [showBulkMenu, setShowBulkMenu] = useState(false);

  const handleResetData = () => {
    if (confirm('Tüm veriler (ithalat, ihracat ve eşleşmeler) kalıcı olarak silinecektir. Emin misiniz?')) {
      setImports([]);
      setExports([]);
      setLinks([]);
      setBulkResults(null);
      setImportFileInfo(null);
      setExportFileInfo(null);
      setPendingExportData(null);
      setPendingExportFileName('');
      setShowBulkMenu(false);
    }
  };

  // Form states
  const [showImportDialog, setShowImportDialog] = useState(false);
  const [showExportDialog, setShowExportDialog] = useState(false);

  // New Import Form
  const [newImport, setNewImport] = useState({
    date: new Date().toISOString().split('T')[0],
    declarationNumber: '',
    gtip: '',
    materialName: 'Dökme Hammadde',
    quantity: 1000,
  });

  // New Export Form
  const [newExport, setNewExport] = useState({
    date: new Date().toISOString().split('T')[0],
    declarationNumber: '',
    gtip: '',
    itemCode: '',
    productName: 'Mamul Ürün',
    quantity: 100,
    consumptionRate: 0.85,
  });

  const handleAddImport = () => {
    const imp: Import = {
      id: crypto.randomUUID(),
      date: newImport.date,
      declarationNumber: newImport.declarationNumber,
      gtip: newImport.gtip,
      materialName: newImport.materialName,
      quantity: newImport.quantity,
      remainingQuantity: newImport.quantity,
    };
    setImports([...imports, imp]);
    setShowImportDialog(false);
  };

  const parseExcelDate = (val: any) => {
    if (!val) return new Date().toISOString().split('T')[0];
    
    // If it's a number (Excel date format)
    if (typeof val === 'number') {
      const d = new Date((val - 25569) * 86400 * 1000);
      return d.toISOString().split('T')[0];
    }

    // If it's a string like DD/MM/YYYY
    if (typeof val === 'string' && val.includes('/')) {
      const parts = val.split('/');
      if (parts.length === 3) {
        const day = parts[0].trim().padStart(2, '0');
        const month = parts[1].trim().padStart(2, '0');
        const year = parts[2].trim();
        const fullYear = year.length === 2 ? `20${year}` : year;
        return `${fullYear}-${month}-${day}`;
      }
    }

    // If it's a string like DD.MM.YYYY
    if (typeof val === 'string' && val.includes('.')) {
      const parts = val.split('.');
      if (parts.length === 3) {
        const day = parts[0].trim().padStart(2, '0');
        const month = parts[1].trim().padStart(2, '0');
        const year = parts[2].trim();
        const fullYear = year.length === 2 ? `20${year}` : year;
        return `${fullYear}-${month}-${day}`;
      }
    }

    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return d.toISOString().split('T')[0];
    }
    
    return new Date().toISOString().split('T')[0];
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      const parseExcelNumber = (val: any): number => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const str = val.toString().trim();
        if (str.includes(',')) {
          return Number(str.replace(/\./g, '').replace(',', '.')) || 0;
        }
        return Number(str) || 0;
      };

      const newImports: Import[] = data.map((row: any) => {
        const findVal = (keys: string[]) => {
          const foundKey = Object.keys(row).find(k => keys.some(search => k.toLowerCase().includes(search.toLowerCase())));
          return foundKey ? row[foundKey] : null;
        };

        const qty = parseExcelNumber(findVal(['miktar', 'qty']));

        return {
          id: crypto.randomUUID(),
          date: parseExcelDate(findVal(['tescil tarihi', 'beyanname tarihi', 'tarih'])),
          declarationNumber: (findVal(['tescil no', 'beyanname no']) || '')?.toString() || '',
          gtip: (findVal(['gtip']) || '0000')?.toString() || '',
          materialName: (findVal(['hammadde', 'ürün']) || 'İthal Hammadde')?.toString(),
          quantity: qty,
          remainingQuantity: qty,
        };
      });

      const totalQty = newImports.reduce((acc, curr) => acc + curr.quantity, 0);
      setImportFileInfo({ name: file.name, rows: newImports.length, total: totalQty });
      setImports(prev => [...prev, ...newImports]);
    };
    reader.readAsBinaryString(file);
  };

  const handleExportListImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      const parseExcelNumber = (val: any): number => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const str = val.toString().trim();
        if (str.includes(',')) {
          return Number(str.replace(/\./g, '').replace(',', '.')) || 0;
        }
        return Number(str) || 0;
      };

      setPendingExportData(data);
      setPendingExportFileName(file.name);
      
      let totalQty = 0;
      data.forEach(row => {
        const findVal = (keys: string[]) => {
          const foundKey = Object.keys(row).find(k => keys.some(search => k.toLowerCase().includes(search.toLowerCase())));
          return foundKey ? row[foundKey] : null;
        };
        totalQty += parseExcelNumber(findVal(['ihracat miktarı', 'miktar']));
      });

      setExportFileInfo({ name: file.name, rows: data.length, total: totalQty });
    };
    reader.readAsBinaryString(file);
  };

  const runAnalysis = () => {
    if (!pendingExportData) {
      alert('Lütfen önce bir ihracat listesi yükleyin.');
      return;
    }
    const data = pendingExportData;
    const fileName = pendingExportFileName;

      const parseExcelNumber = (val: any): number => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const str = val.toString().trim();
        if (str.includes(',')) {
          return Number(str.replace(/\./g, '').replace(',', '.')) || 0;
        }
        return Number(str) || 0;
      };

      let currentImports = [...imports];
      let currentExports = [...exports];
      let currentLinks = [...links];
      const analysisItems: any[] = [];
      let totalExportQty = 0;

      data.forEach((row: any) => {
        const findVal = (keys: string[]) => {
          const foundKey = Object.keys(row).find(k => keys.some(search => k.toLowerCase().includes(search.toLowerCase())));
          return foundKey ? row[foundKey] : null;
        };

        const itemCode = (findVal(['satır kodu', 'kalem kodu']) || '')?.toString() || '';
        let rate = 0.85; 
        if (itemCode === '26.1.01378.001') rate = 0.928;
        else if (itemCode === '26.1.01378.002') rate = 0.865;

        const expDate = parseExcelDate(findVal(['tescil tarihi', 'beyanname tarihi', 'tarih']));
        const qty = parseExcelNumber(findVal(['ihracat miktarı', 'miktar']));
        totalExportQty += qty;
        const decNo = (findVal(['tescil no', 'beyanname no']) || '')?.toString() || '';

        // Pre-check for available stock before date
        const requiredQty = qty * rate;
        const availableBeforeDate = currentImports.filter(imp => {
          const importDate = new Date(imp.date);
          const exportDate = new Date(expDate);
          const timeDiff = exportDate.getTime() - importDate.getTime();
          const daysDiff = timeDiff / (1000 * 3600 * 24);
          return daysDiff >= 1 && imp.remainingQuantity > 0;
        }).reduce((acc, curr) => acc + curr.remainingQuantity, 0);

        if (availableBeforeDate < requiredQty - 0.001) { // Floating point safety
          analysisItems.push({
            id: crypto.randomUUID(),
            declarationNumber: decNo,
            date: expDate,
            quantity: qty,
            consumptionRate: rate,
            itemCode: itemCode,
            status: 'FAIL',
            error: `Stok yetersiz veya üretim tarihi hatalı. (Gereken: ${requiredQty.toLocaleString()} KG, Mevcut: ${availableBeforeDate.toLocaleString()} KG)`
          });
          return;
        }

        const exp: Export = {
          id: crypto.randomUUID(),
          date: expDate,
          declarationNumber: decNo,
          gtip: (findVal(['gtip']) || '0000')?.toString(),
          itemCode: itemCode,
          productName: (findVal(['ürün', 'mamul']) || 'Excel İhracatı')?.toString(),
          quantity: qty,
          consumptionRate: rate,
          status: Status.PENDING,
        };

        const { updatedImports, links: newLinks, status } = matchExportWithImports(exp, currentImports);

        if (status === Status.PENDING) {
          analysisItems.push({
            id: exp.id,
            declarationNumber: decNo,
            date: expDate,
            quantity: qty,
            consumptionRate: rate,
            itemCode: itemCode,
            status: 'FAIL',
            error: 'FIFO eşleşmesi sırasında stok yetersiz kaldı.'
          });
        } else {
          currentImports = updatedImports;
          currentExports.push({ ...exp, status });
          currentLinks = [...currentLinks, ...newLinks];
          analysisItems.push({
            id: exp.id,
            declarationNumber: decNo,
            date: expDate,
            quantity: qty,
            consumptionRate: rate,
            itemCode: itemCode,
            status: status === Status.PARTIAL ? 'FAIL' : 'OK',
            error: status === Status.PARTIAL ? 'Stok sadece kısmen karşılanabildi.' : undefined
          });
        }
      });

      setImportFileInfo({ ...importFileInfo!, rows: imports.length, total: imports.reduce((a, b) => a + b.quantity, 0) });
      setExportFileInfo({ name: fileName, rows: data.length, total: totalExportQty });

      setImports(currentImports);
      setExports(currentExports);
      setLinks(currentLinks);
      
      const failedCount = analysisItems.filter(i => i.status === 'FAIL').length;
      setBulkResults({
        success: failedCount === 0,
        message: failedCount > 0 
          ? `${analysisItems.length} işlemden ${failedCount} tanesi stok yetersizliği veya tarih hatası nedeniyle TAMAMLANAMADI.` 
          : "Tüm Excel kayıtları başarıyla işlendi ve stoktan düşüldü.",
        items: analysisItems.map(item => ({
          ...item,
          linkedImports: currentLinks
            .filter(l => l.exportId === item.id)
            .map(l => {
              const imp = currentImports.find(i => i.id === l.importId);
              return imp ? `${imp.declarationNumber} (${l.consumedQuantity.toLocaleString()} KG)` : undefined;
            }).filter(Boolean) as string[]
        }))
      });
  };

  const downloadImportTemplate = () => {
    const data = [
      {
        'Beyanname No': '24063400IM000001',
        'Tarih': '01.01.2024',
        'Miktar': 1000,
        'GTIP': '260111000000',
        'Hammadde': 'Demir Cevheri'
      }
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 20}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "İthalat Şablonu");
    XLSX.writeFile(wb, "Ithalat_Yukleme_Sablonu.xlsx");
  };

  const downloadExportTemplate = () => {
    const data = [
      {
        'Beyanname No': '24063400EX000001',
        'Tarih': '15.01.2024',
        'Satır Kodu': '26.1.01378.001',
        'Miktar': 100,
        'GTIP': '720110110000',
        'Ürün': 'Pik Demir'
      }
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 20}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "İhracat Şablonu");
    XLSX.writeFile(wb, "Ihracat_Yukleme_Sablonu.xlsx");
  };

  const downloadExportReport = () => {
    // If we are currently looking at a bulk result, download that specific analysis
    // Otherwise, download the entire successfully added export list.
    const sourceData = bulkResults ? bulkResults.items : exports.map(e => ({
      declarationNumber: e.declarationNumber,
      date: e.date,
      itemCode: e.itemCode,
      quantity: e.quantity,
      consumptionRate: e.consumptionRate,
      status: e.status === Status.COMPLETED ? 'OK' : 'FAIL',
      id: e.id,
      error: e.status === Status.PARTIAL ? 'Kısmi karşılandı' : undefined
    }));

    const reportData: any[] = [];
    
    sourceData.forEach(item => {
      const exportLinks = links.filter(l => l.exportId === item.id);
      
      if (exportLinks.length > 0) {
        // If there are links, create a row for each link (especially for split exports)
        exportLinks.forEach(link => {
          const imp = imports.find(i => i.id === link.importId);
          reportData.push({
            'Beyanname No': item.declarationNumber,
            'Beyanname Tarihi': new Date(item.date).toLocaleDateString('tr-TR'),
            'Satır Kodu': item.itemCode || '-',
            'Toplam İhracat Miktarı': item.quantity,
            'Düşülen Stok (KG)': link.consumedQuantity,
            'İlgili İthalat': imp ? imp.declarationNumber : 'BELİRSİZ',
            'Parçalı mı?': exportLinks.length > 1 ? 'EVET' : 'HAYIR',
            'Durum': 'BAŞARILI',
            'Analiz Notu': exportLinks.length > 1 ? 'Parçalı FIFO düşümü yapıldı' : ''
          });
        });
      } else {
        // If no links (FAIL or PARTIAL with no first match), show failure row
        reportData.push({
          'Beyanname No': item.declarationNumber,
          'Beyanname Tarihi': new Date(item.date).toLocaleDateString('tr-TR'),
          'Satır Kodu': item.itemCode || '-',
          'Toplam İhracat Miktarı': item.quantity,
          'Düşülen Stok (KG)': 0,
          'İlgili İthalat': 'YOK / HATA',
          'Parçalı mı?': 'HAYIR',
          'Durum': 'HATA',
          'Analiz Notu': item.error || 'Stok bulunamadı'
        });
      }
    });

    const ws = XLSX.utils.json_to_sheet(reportData);
    
    // Auto-size columns logic
    const wscols = [
      {wch: 20}, // No
      {wch: 15}, // Tarih
      {wch: 15}, // Kod
      {wch: 25}, // Toplam Miktar
      {wch: 20}, // Stok
      {wch: 25}, // İthalat
      {wch: 15}, // Parçalı
      {wch: 15}, // Durum
      {wch: 50}, // Not
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Analiz Raporu");

    // Add Imports sheet
    const importData = imports.map(imp => ({
      'Beyanname No': imp.declarationNumber,
      'Tarih': new Date(imp.date).toLocaleDateString('tr-TR'),
      'Giriş Miktarı': imp.quantity,
      'Kalan Miktar': imp.remainingQuantity,
      'Durum': imp.remainingQuantity < 0.001 ? 'TÜKENDİ' : 'AKTİF'
    }));
    const wsImports = XLSX.utils.json_to_sheet(importData);
    wsImports['!cols'] = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, wsImports, "İthalat Listesi");

    // Add Exports sheet
    const exportData = exports.map(exp => ({
      'Beyanname No': exp.declarationNumber,
      'Tarih': new Date(exp.date).toLocaleDateString('tr-TR'),
      'Satır Kodu': exp.itemCode,
      'İhracat Miktarı': exp.quantity,
      'Sarfiyat Oranı': exp.consumptionRate,
      'Hesaplanan Stok (KG)': exp.quantity * exp.consumptionRate,
      'Analiz Durumu': exp.status === Status.COMPLETED ? 'BAŞARILI' : (exp.status === Status.PARTIAL ? 'KISMI' : 'HATA')
    }));
    const wsExports = XLSX.utils.json_to_sheet(exportData);
    wsExports['!cols'] = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 20}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, wsExports, "İhracat Listesi");

    XLSX.writeFile(wb, `FIFO_Stok_Takip_Analyzi_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleAddExport = () => {
    const totalToConsume = newExport.quantity * newExport.consumptionRate;
    
    // Check if there is ANY valid import before this date
    const availableBeforeDate = imports.filter(imp => {
      const importDate = new Date(imp.date);
      const exportDate = new Date(newExport.date);
      const daysDiff = (exportDate.getTime() - importDate.getTime()) / (1000 * 3600 * 24);
      return daysDiff >= 1 && imp.remainingQuantity > 0;
    }).reduce((acc, curr) => acc + curr.remainingQuantity, 0);

    if (availableBeforeDate <= 0) {
      alert("Hata: Seçilen tarihten önce yapılmış ve bakiyesi olan bir ithalat kaydı bulunamadı. Rejim gereği önceden ithalat şartı aranmaktadır.");
      return;
    }

    const expId = crypto.randomUUID();
    const exp: Export = {
      id: expId,
      date: newExport.date,
      declarationNumber: newExport.declarationNumber,
      gtip: newExport.gtip,
      itemCode: newExport.itemCode,
      productName: newExport.productName,
      quantity: newExport.quantity,
      consumptionRate: newExport.consumptionRate,
      status: Status.PENDING,
    };

    const { updatedImports, links: newLinks, status } = matchExportWithImports(exp, imports);
    
    if (status === Status.PENDING) {
      alert("Hata: Mevcut stoklar bu sarfiyatı karşılamıyor. Lütfen miktarı veya tarihi kontrol edin.");
      return;
    }

    setImports(updatedImports);
    setExports([...exports, { ...exp, status }]);
    setLinks([...links, ...newLinks]);
    setShowExportDialog(false);
  };

  const availableStockForSelectedDate = useMemo(() => {
    return imports.filter(imp => {
      const importDate = new Date(imp.date);
      const exportDate = new Date(newExport.date);
      const daysDiff = (exportDate.getTime() - importDate.getTime()) / (1000 * 3600 * 24);
      return daysDiff >= 1;
    }).reduce((acc, curr) => acc + curr.remainingQuantity, 0);
  }, [imports, newExport.date]);

  const totalImportStock = useMemo(() => 
    imports.reduce((acc, curr) => acc + curr.remainingQuantity, 0), 
    [imports]
  );

  const totalRawImport = useMemo(() => 
    imports.reduce((acc, curr) => acc + curr.quantity, 0), 
    [imports]
  );

  return (
    <div className="flex h-screen bg-[#f8fafc] font-sans text-slate-900 overflow-hidden">
      {/* Sidebar */}
      <aside className="w-64 bg-[#111827] flex flex-col shrink-0">
        <div className="p-8 pb-10 flex items-center gap-3">
          <div className="w-9 h-9 bg-[#10b981] rounded-lg flex items-center justify-center text-white font-black text-sm">DIR</div>
          <div>
            <h1 className="text-sm font-black text-white tracking-tight leading-none">DIR Stok Takip</h1>
            <p className="text-[9px] text-slate-500 font-medium mt-1">Dahilde İşleme Rejimi</p>
          </div>
        </div>

        <nav className="flex-1 px-4 space-y-1">
          {[
            { id: 'overview', label: 'Özet', icon: LayoutGrid },
            { id: 'links', label: 'Eşleşmeler', icon: Link2 },
            { id: 'stock', label: 'Stok', icon: Box },
            { id: 'errors', label: 'Hatalar', icon: AlertTriangle },
          ].map(item => (
            <button
              key={item.id}
              onClick={() => setActiveMenu(item.id as any)}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition-all ${
                activeMenu === item.id 
                  ? 'bg-slate-800/100 text-white shadow-lg' 
                  : 'text-slate-500 hover:text-slate-300 hover:bg-slate-800/50'
              }`}
            >
              <item.icon size={18} className={activeMenu === item.id ? 'text-white' : 'text-slate-500'} />
              {item.label}
            </button>
          ))}
        </nav>

        <div className="p-6">
          <div className="p-4 bg-slate-800/30 rounded-2xl border border-slate-800/50">
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-widest mb-3">SARFİYAT ORANLARI</p>
            <div className="space-y-2">
              <div className="flex justify-between text-[10px] text-slate-400">
                <span>26.1...001:</span>
                <span className="font-bold text-[#10b981]">0.928</span>
              </div>
              <div className="flex justify-between text-[10px] text-slate-400">
                <span>26.1...002:</span>
                <span className="font-bold text-[#10b981]">0.865</span>
              </div>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Top Header */}
        <header className="bg-white px-10 py-6 border-b border-slate-100 shrink-0">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em] mb-1">EXCEL İŞLEMLERİ</p>
              <h2 className="text-[26px] font-black text-slate-900 tracking-tight leading-tight">FIFO ithalat-ihracat analizi</h2>
            </div>

            <div className="flex items-center gap-2">
              <button 
                onClick={() => setShowImportDialog(true)}
                className="flex items-center gap-2 bg-[#10b981] text-white px-5 py-2.5 rounded-lg font-black text-xs hover:bg-[#059669] transition-all shadow-sm shadow-emerald-100"
              >
                <Plus size={14} />
                Manuel İthalat
              </button>

              <button 
                onClick={() => setShowExportDialog(true)}
                className="flex items-center gap-2 bg-[#1d4d46] text-white px-5 py-2.5 rounded-lg font-black text-xs hover:bg-[#153a35] transition-all shadow-sm shadow-slate-200"
              >
                <Plus size={14} />
                Manuel İhracat
              </button>

              <div className="w-px h-8 bg-slate-200 mx-2" />

              <div className="flex flex-col gap-1">
                <label className="flex items-center gap-2 bg-white border border-slate-200 text-slate-900 px-5 py-2.5 rounded-lg font-black text-xs hover:bg-slate-50 transition-all shadow-sm cursor-pointer">
                  <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleExcelImport} />
                  <ArrowDownCircle size={14} className="text-slate-400" />
                  Excel İthalat
                </label>
                <button 
                  onClick={downloadImportTemplate}
                  className="text-[9px] font-black text-emerald-600 hover:text-emerald-700 text-left px-1 flex items-center gap-1 transition-colors"
                >
                  <FileSpreadsheet size={10} />
                  Şablon İndir
                </button>
              </div>

              <div className="flex flex-col gap-1">
                <label className="flex items-center gap-2 bg-white border border-slate-200 text-slate-900 px-5 py-2.5 rounded-lg font-black text-xs hover:bg-slate-50 transition-all shadow-sm cursor-pointer">
                  <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleExportListImport} />
                  <ArrowUpCircle size={14} className="text-slate-400" />
                  Excel İhracat
                </label>
                <button 
                  onClick={downloadExportTemplate}
                  className="text-[9px] font-black text-[#1d4d46] hover:text-slate-900 text-left px-1 flex items-center gap-1 transition-colors"
                >
                  <FileSpreadsheet size={10} />
                  Şablon İndir
                </button>
              </div>

              <button 
                onClick={runAnalysis}
                disabled={!pendingExportData}
                className={`flex items-center gap-2 px-6 py-2.5 rounded-lg font-black text-xs transition-all shadow-[0_4px_12px_rgba(29,77,70,0.25)] ${
                  !pendingExportData 
                  ? 'bg-slate-100 text-slate-400 cursor-not-allowed shadow-none' 
                  : 'bg-[#1d4d46] text-white hover:bg-[#153a35]'
                }`}
              >
                <TrendingUp size={14} />
                Analizi Çalıştır
              </button>

              <button 
                onClick={handleResetData}
                className="p-2.5 rounded-lg border border-rose-100 text-rose-300 hover:bg-rose-50 hover:text-rose-500 transition-all"
              >
                <RotateCcw size={14} />
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 mt-8">
            <div className={`p-5 rounded-2xl border flex items-center gap-5 transition-all ${importFileInfo ? 'bg-[#f0fdf9] border-[#dcfce7]' : 'bg-slate-50 border-slate-100'}`}>
              <div className={`w-12 h-12 rounded-xl flex items-center justify-center shadow-sm ${importFileInfo ? 'bg-[#d1fae5] text-[#10b981]' : 'bg-slate-200 text-slate-400'}`}>
                <Box size={22} />
              </div>
              <div className="flex-1">
                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest leading-none mb-1.5">İTHALAT LİSTESİ</p>
                <div className="flex items-center justify-between">
                  <p className="text-sm font-black text-slate-800 leading-none truncate max-w-[180px]">{importFileInfo?.name || 'ithalat.xlsx'}</p>
                </div>
                {importFileInfo && (
                  <p className="text-[10px] text-slate-400 font-bold mt-1.5">{importFileInfo.rows} satır, toplam {importFileInfo.total.toLocaleString()} miktar</p>
                )}
              </div>
            </div>

            <div className={`p-5 rounded-2xl border flex items-center gap-5 transition-all ${exportFileInfo ? 'bg-[#f0fdf9] border-[#dcfce7]' : 'bg-slate-50 border-slate-100'}`}>
              <div className={`w-12 h-12 rounded-xl flex items-center justify-center shadow-sm ${exportFileInfo ? 'bg-[#d1fae5] text-[#10b981]' : 'bg-slate-200 text-slate-400'}`}>
                <FileText size={22} />
              </div>
              <div className="flex-1">
                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest leading-none mb-1.5">İHRACAT LİSTESİ</p>
                <div className="flex items-center justify-between">
                  <p className="text-sm font-black text-slate-800 leading-none truncate max-w-[180px]">{exportFileInfo?.name || 'ihracat.xlsx'}</p>
                </div>
                {exportFileInfo && (
                  <p className="text-[10px] text-slate-400 font-bold mt-1.5">{exportFileInfo.rows} satır, toplam {exportFileInfo.total.toLocaleString()} ihracat miktarı</p>
                )}
              </div>
            </div>
          </div>
        </header>

        {/* Scrollable Content Area */}
        <main className="flex-1 overflow-y-auto p-10 space-y-10 scrollbar-hide">
          {/* Key Stat Cards */}
          <div className="grid grid-cols-4 gap-6">
            {[
              { label: 'TOPLAM İTHALAT', value: totalRawImport.toLocaleString(), sub: 'Miktar' },
              { label: 'KALAN STOK', value: totalImportStock.toLocaleString(), sub: 'FIFO sonrası' },
              { label: 'İHRACAT KAYDI', value: exports.length, sub: 'Analize giren' },
              { label: 'UYARI', value: bulkResults?.items.filter(i => i.status === 'FAIL').length || 0, sub: 'Sorunlu veya kısmi' },
            ].map(card => (
              <div key={card.label} className="bg-white p-7 rounded-2xl border border-slate-100 shadow-sm hover:shadow-md transition-all">
                <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.1em] leading-none mb-5">{card.label}</p>
                <p className="text-[32px] font-black text-slate-900 tracking-tighter tabular-nums leading-none mb-2">{card.value}</p>
                <p className="text-[10px] font-bold text-slate-400 italic leading-none">{card.sub}</p>
              </div>
            ))}
          </div>

          <div className="bg-white rounded-[32px] border border-slate-100 shadow-2xl shadow-slate-200/50 overflow-hidden flex flex-col min-h-[600px]">
            <div className="p-8 border-b border-slate-100 flex items-center justify-between shrink-0">
              <div className="flex p-1.5 bg-slate-50 rounded-2xl items-center gap-1">
                {[
                  { id: 'overview', label: 'Özet', count: null },
                  { id: 'links', label: 'Eşleşmeler', count: links.length },
                  { id: 'stock', label: 'Stok', count: imports.filter(i => i.remainingQuantity > 0).length },
                  { id: 'errors', label: 'Hatalar', count: bulkResults?.items.filter(i => i.status === 'FAIL').length || 0 },
                ].map(tab => (
                  <button
                    key={tab.id}
                    onClick={() => setActiveMenu(tab.id as any)}
                    className={`flex items-center gap-2.5 px-6 py-2.5 rounded-xl font-black text-[11px] transition-all uppercase tracking-tight ${
                      activeMenu === tab.id 
                        ? 'bg-[#eefcf7] text-[#10b981] shadow-sm' 
                         : 'text-slate-400 hover:text-slate-600'
                    }`}
                  >
                    {tab.label}
                    {tab.count !== null && (
                      <span className={`px-2 py-0.5 rounded-lg text-[10px] font-bold ${
                        activeMenu === tab.id ? 'bg-[#d1fae5] text-[#10b981]' : 'bg-slate-200/50 text-slate-400'
                      }`}>{tab.count}</span>
                    )}
                  </button>
                ))}
              </div>

              <div className="flex items-center gap-3">
                <div className="relative group">
                  <Search size={15} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-slate-500 transition-colors" />
                  <input 
                    type="text" 
                    placeholder="Beyanname veya ürün kodu ara"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="bg-[#f0f4f8] border-none rounded-xl pl-11 pr-5 py-3 text-[11px] font-black w-72 focus:ring-2 focus:ring-slate-200 transition-all placeholder:text-slate-400 uppercase tracking-tight"
                  />
                </div>
                <button 
                  onClick={downloadExportReport}
                  className="bg-[#1d4d46] text-white px-7 py-3 rounded-xl font-black text-[11px] uppercase hover:bg-[#153a35] transition-all shadow-lg flex items-center gap-2.5"
                >
                  <FileSpreadsheet size={16} />
                  Excel Olarak İndir
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-x-auto">
              {activeMenu === 'stock' && (
                <>
                  <table className="w-full text-left">
                    <thead className="bg-[#f8fafc]/50 border-b border-slate-100">
                      <tr>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest">İTHALAT BEYANNAME</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest">TARİH</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">GİRİŞ MİKTARI</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">KULLANILAN</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">KALAN</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-center">DURUM</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {imports.filter(i => 
                        i.declarationNumber.toLowerCase().includes(searchTerm.toLowerCase()) ||
                        i.materialName.toLowerCase().includes(searchTerm.toLowerCase())
                      ).map(imp => {
                        const usedQty = imp.quantity - imp.remainingQuantity;
                        const isConsumed = imp.remainingQuantity < 0.001;

                        return (
                          <tr key={imp.id} className="hover:bg-[#f8fafc] transition-colors group">
                            <td className="px-10 py-5.5">
                              <p className="text-xs font-bold text-slate-700 tracking-tight">{imp.declarationNumber}</p>
                            </td>
                            <td className="px-10 py-5.5">
                              <p className="text-xs font-bold text-slate-500 tracking-tight">{new Date(imp.date).toLocaleDateString('tr-TR')}</p>
                            </td>
                            <td className="px-10 py-5.5 text-right">
                              <p className="text-xs font-bold text-slate-400 tabular-nums">{imp.quantity.toLocaleString('tr-TR')}</p>
                            </td>
                            <td className="px-10 py-5.5 text-right">
                              <p className="text-xs font-black text-slate-500 tabular-nums">{usedQty.toLocaleString('tr-TR')}</p>
                            </td>
                            <td className="px-10 py-5.5 text-right">
                              <p className="text-xs font-black text-slate-800 tabular-nums">{imp.remainingQuantity.toLocaleString('tr-TR', { minimumFractionDigits: 2 })}</p>
                            </td>
                            <td className="px-10 py-5.5 text-center">
                              {isConsumed ? (
                                <span className="bg-slate-100 text-slate-400 px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-[0.15em] border border-slate-200/50">TÜKENDİ</span>
                              ) : (
                                <span className="bg-[#eefcf7] text-[#10b981] px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-[0.15em] border border-[#dcfce7]">AKTİF</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {imports.length === 0 && (
                    <div className="flex flex-col items-center justify-center py-32 text-center space-y-5">
                      <div className="w-20 h-20 bg-slate-50 rounded-3xl flex items-center justify-center text-slate-200 border border-slate-100">
                        <Box size={40} />
                      </div>
                      <div>
                        <h3 className="text-xl font-black text-slate-900 tracking-tight">Stok kaydı yok</h3>
                        <p className="text-xs text-slate-400 font-bold mt-1.5 max-w-[280px] leading-relaxed uppercase tracking-tight">Seçili bölümde gösterilecek kayıt bulunmuyor.</p>
                      </div>
                    </div>
                  )}
                </>
              )}

              {activeMenu === 'links' && (
                <>
                  <table className="w-full text-left">
                    <thead className="bg-[#f8fafc]/50 border-b border-slate-100">
                      <tr>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest">İHRACAT BEYANNAME</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest">İTHALAT KAYNAĞI</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">İHRACAT MİKTAR</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">DÜŞÜLEN STOK</th>
                        <th className="px-10 py-6 text-[9px] font-black text-slate-400 uppercase tracking-widest text-center">DURUM</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {exports.filter(exp => 
                        exp.declarationNumber.toLowerCase().includes(searchTerm.toLowerCase()) ||
                        exp.itemCode.toLowerCase().includes(searchTerm.toLowerCase())
                      ).flatMap(exp => {
                        const exportLinks = links.filter(l => l.exportId === exp.id);
                        
                        // If there are no successful links, show the failure row
                        if (exportLinks.length === 0) {
                          return [(
                            <tr key={exp.id} className="hover:bg-[#f8fafc] transition-colors group">
                              <td className="px-10 py-5.5">
                                <p className="text-xs font-bold text-slate-700 tracking-tight">{exp.declarationNumber}</p>
                                <p className="text-[10px] text-slate-400 font-medium mt-0.5">{new Date(exp.date).toLocaleDateString('tr-TR')}</p>
                              </td>
                              <td className="px-10 py-5.5 text-[9px] text-rose-500 font-black uppercase italic">
                                Stok Bulunamadı
                              </td>
                              <td className="px-10 py-5.5 text-right text-xs font-bold text-slate-500 tabular-nums">
                                {exp.quantity.toLocaleString()} AD
                              </td>
                              <td className="px-10 py-5.5 text-right text-xs font-black text-rose-400 tabular-nums">
                                0,00 KG
                              </td>
                              <td className="px-10 py-5.5 text-center">
                                <ThemeStatusBadge status={exp.status} />
                              </td>
                            </tr>
                          )];
                        }

                        const isSplit = exportLinks.length > 1;

                        // Create a row for each link
                        return exportLinks.map((link, idx) => {
                          const imp = imports.find(i => i.id === link.importId);
                          return (
                            <tr 
                              key={`${exp.id}-${idx}`} 
                              className={`transition-colors group ${isSplit ? 'bg-red-600/10 hover:bg-red-600/20' : 'hover:bg-[#f8fafc]'}`}
                            >
                              <td className="px-10 py-5.5">
                                <div className="flex items-center gap-2">
                                  {isSplit && <div className="w-1.5 h-6 bg-red-600 rounded-full shrink-0" />}
                                  <div>
                                    <p className={`text-xs font-bold tracking-tight ${isSplit ? 'text-red-900' : 'text-slate-700'}`}>{exp.declarationNumber}</p>
                                    <p className={`text-[10px] font-medium mt-0.5 ${isSplit ? 'text-red-400' : 'text-slate-400'}`}>{new Date(exp.date).toLocaleDateString('tr-TR')}</p>
                                  </div>
                                </div>
                              </td>
                              <td className="px-10 py-5.5">
                                <span className={`px-2.5 py-1 rounded-md text-[10px] font-black border transition-all ${isSplit ? 'bg-red-600 text-white border-red-500' : 'bg-blue-50 text-blue-600 border-blue-100'}`}>
                                  {imp ? imp.declarationNumber : 'Eşleşiyor...'}
                                </span>
                              </td>
                              <td className={`px-10 py-5.5 text-right text-xs font-bold tabular-nums ${isSplit ? 'text-red-900' : 'text-slate-500'}`}>
                                {exp.quantity.toLocaleString()} AD
                              </td>
                              <td className={`px-10 py-5.5 text-right text-xs font-black tabular-nums transition-colors ${isSplit ? 'bg-red-600 text-white rounded-lg shadow-sm' : 'text-slate-700'}`}>
                                {link.consumedQuantity.toLocaleString('tr-TR', { minimumFractionDigits: 2 })} KG
                              </td>
                              <td className="px-10 py-5.5 text-center">
                                <ThemeStatusBadge status={isSplit ? Status.COMPLETED : exp.status} />
                              </td>
                            </tr>
                          );
                        });
                      })}
                    </tbody>
                  </table>
                  {exports.length === 0 && (
                    <div className="flex flex-col items-center justify-center py-32 text-center space-y-5">
                      <div className="w-20 h-20 bg-slate-50 rounded-3xl flex items-center justify-center text-slate-200 border border-slate-100">
                        <Link2 size={40} />
                      </div>
                      <div>
                        <h3 className="text-xl font-black text-slate-900 tracking-tight">Eşleşme kaydı yok</h3>
                        <p className="text-xs text-slate-400 font-bold mt-1.5 max-w-[280px] leading-relaxed uppercase tracking-tight">Analiz yapılmadı veya eşleşen kayıt bulunmuyor.</p>
                      </div>
                    </div>
                  )}
                </>
              )}

              {activeMenu === 'errors' && (
                <>
                  <table className="w-full text-left">
                    <thead className="bg-[#fef2f2]/50 border-b border-rose-100">
                      <tr>
                        <th className="px-10 py-6 text-[9px] font-black text-rose-400 uppercase tracking-widest">SORUNLU BEYANNAME</th>
                        <th className="px-10 py-6 text-[9px] font-black text-rose-400 uppercase tracking-widest">HATA DETAYI</th>
                        <th className="px-10 py-6 text-[9px] font-black text-rose-400 uppercase tracking-widest text-right">MİKTAR</th>
                        <th className="px-10 py-6 text-[9px] font-black text-rose-400 uppercase tracking-widest text-center">İŞLEM</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-rose-50">
                      {(bulkResults?.items || []).filter(i => i.status === 'FAIL').map((item, idx) => (
                        <tr key={idx} className="hover:bg-rose-50/50 transition-colors group">
                          <td className="px-10 py-5.5">
                            <p className="text-xs font-bold text-rose-900 tracking-tight">{item.declarationNumber}</p>
                            <p className="text-[10px] text-rose-400 font-medium mt-0.5">{new Date(item.date).toLocaleDateString('tr-TR')}</p>
                          </td>
                          <td className="px-10 py-5.5">
                            <p className="text-[11px] font-medium text-rose-600 leading-snug max-w-sm italic">
                              {item.error}
                            </p>
                          </td>
                          <td className="px-10 py-5.5 text-right text-xs font-black text-rose-900 tabular-nums">
                            {item.quantity.toLocaleString()} AD
                          </td>
                          <td className="px-10 py-5.5 text-center">
                            <span className="bg-rose-100 text-rose-600 px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest">REKETELİ</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {(bulkResults?.items.filter(i => i.status === 'FAIL').length === 0) && (
                    <div className="flex flex-col items-center justify-center py-32 text-center space-y-5">
                      <div className="w-20 h-20 bg-emerald-50 rounded-3xl flex items-center justify-center text-emerald-200 border border-emerald-100">
                        <CheckCircle2 size={40} />
                      </div>
                      <div>
                        <h3 className="text-xl font-black text-emerald-900 tracking-tight">Hata saptanmadı</h3>
                        <p className="text-xs text-emerald-400 font-bold mt-1.5 max-w-[280px] leading-relaxed uppercase tracking-tight">Tüm kayıtlar başarıyla analiz edildi.</p>
                      </div>
                    </div>
                  )}
                </>
              )}

              {activeMenu === 'overview' && (
                <div className="p-10 space-y-10">
                  <div className="grid grid-cols-2 gap-8">
                    <div className="space-y-4">
                      <h3 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                        <div className="w-2 h-4 bg-emerald-500 rounded-sm"></div>
                        HIZLI ÖZET
                      </h3>
                      <div className="grid grid-cols-2 gap-4">
                        <div className="p-5 bg-slate-50 rounded-2xl border border-slate-100">
                          <p className="text-[9px] font-black text-slate-400 uppercase mb-2">Doluluk Oranı</p>
                          <p className="text-2xl font-black text-slate-800">
                            {totalRawImport > 0 ? Math.round(((totalRawImport - totalImportStock) / totalRawImport) * 100) : 0}%
                          </p>
                        </div>
                        <div className="p-5 bg-slate-50 rounded-2xl border border-slate-100">
                          <p className="text-[9px] font-black text-slate-400 uppercase mb-2">Başarı Oranı</p>
                          <p className="text-2xl font-black text-[#10b981]">
                            {exports.length > 0 ? Math.round((exports.filter(e => e.status === Status.COMPLETED).length / exports.length) * 100) : 100}%
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="space-y-4">
                      <h3 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                        <div className="w-2 h-4 bg-blue-500 rounded-sm"></div>
                        SİSTEM BİLGİSİ
                      </h3>
                      <div className="p-6 bg-blue-50 rounded-[28px] border border-blue-100 flex items-center gap-5">
                        <div className="w-12 h-12 bg-white rounded-xl flex items-center justify-center text-blue-500 shadow-sm">
                          <TrendingUp size={24} />
                        </div>
                        <div>
                          <p className="text-xs font-black text-blue-900 leading-tight">Analiz Motoru v1.0.4</p>
                          <p className="text-[10px] text-blue-600 font-medium mt-1 uppercase tracking-widest italic">FIFO Modu: Aktif / Kesin Eşleşme</p>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="p-8 bg-slate-900 rounded-[32px] text-white flex items-center justify-between shadow-xl shadow-slate-900/20">
                    <div className="space-y-1">
                      <h4 className="text-lg font-black tracking-tight italic">Yeni İhracat Analizi Zamanı</h4>
                      <p className="text-xs text-slate-400 font-medium">Excel dosyanızı yükleyin ve sistemin sihrini göstermesine izin verin.</p>
                    </div>
                    <label className="bg-[#10b981] text-white px-8 py-3 rounded-xl font-black text-[11px] uppercase tracking-widest hover:bg-[#059669] transition-all cursor-pointer shadow-lg shadow-[#10b981]/20">
                      <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleExportListImport} />
                      Dosya Seç
                    </label>
                  </div>
                </div>
              )}
            </div>
          </div>
        </main>
      </div>


      {/* Legacy Dialogs kept for functionality support */}
      <AnimatePresence>
        {bulkResults && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="card bg-white w-full max-w-2xl max-h-[80vh] overflow-hidden flex flex-col"
            >
              <div className={`p-6 flex items-center justify-between border-b ${bulkResults.success ? 'bg-emerald-50 border-emerald-100' : 'bg-rose-50 border-rose-100'}`}>
                <div className="flex items-center gap-3">
                  {bulkResults.success ? (
                    <CheckCircle2 className="text-emerald-500" size={24} />
                  ) : (
                    <AlertCircle className="text-rose-500" size={24} />
                  )}
                  <div>
                    <h3 className={`text-lg font-black uppercase tracking-tighter ${bulkResults.success ? 'text-emerald-800' : 'text-rose-800'}`}>
                      {bulkResults.success ? 'Analiz Tamamlandı' : 'Analiz Hatası / Eksik Stok'}
                    </h3>
                    <p className={`text-[10px] font-bold ${bulkResults.success ? 'text-emerald-600' : 'text-rose-600'}`}>
                      TOPLU İŞLEM RAPORU
                    </p>
                  </div>
                </div>
                <button 
                  onClick={() => setBulkResults(null)}
                  className="p-2 hover:bg-white/50 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
              
              <div className="flex-1 overflow-y-auto p-6 space-y-6">
                <p className="text-sm font-semibold text-slate-600 leading-relaxed">
                  {bulkResults.message}
                </p>

                <div className="space-y-3">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest italic">İşlem Detayları</h4>
                  <div className="border border-slate-100 rounded-xl overflow-hidden">
                    <table className="w-full text-left">
                      <thead className="bg-slate-50">
                        <tr>
                          <th className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-tighter">Beyanname No</th>
                          <th className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-tighter">Durum</th>
                          <th className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-tighter">Analiz Notu</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {bulkResults.items.map((item: any, idx: number) => (
                          <tr key={idx} className={item.status === 'FAIL' ? 'bg-rose-50/30' : ''}>
                            <td className="px-4 py-3 text-xs font-bold text-slate-700">
                              <div>{item.declarationNumber}</div>
                              {item.linkedImports && item.linkedImports.length > 0 && (
                                <div className="mt-1 flex flex-wrap gap-1">
                                  {item.linkedImports.map((li: string, lidx: number) => (
                                    <span key={lidx} className="bg-blue-50 text-blue-600 px-1 py-0.5 rounded-[4px] text-[8px] font-black uppercase tracking-tighter border border-blue-100 italic">
                                      {li}
                                    </span>
                                  ))}
                                </div>
                              )}
                            </td>
                            <td className="px-4 py-3">
                              <span className={`text-[10px] font-black uppercase px-2 py-0.5 rounded ${item.status === 'OK' ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'}`}>
                                {item.status === 'OK' ? 'BAŞARILI' : 'SORUNLU'}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-[10px] font-medium text-slate-500 italic">
                              {item.error || 'Stok yeterli, FIFO eşleşmesi yapıldı.'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {!bulkResults.success && (
                  <div className="p-4 bg-amber-50 rounded-xl border border-amber-100 flex gap-3">
                    <AlertCircle className="text-amber-500 shrink-0" size={18} />
                    <div>
                      <p className="text-xs font-black text-amber-800 uppercase tracking-tighter mb-1">Mantar / Mantık Analizi</p>
                      <p className="text-xs text-amber-700 leading-relaxed italic">
                        DIR rejimi gereği ihracattan önce stokta hammadde bulunmalıdır. Başarısız olan satırlar için ya tarih hatası vardır (ithalat ihracattan sonra) ya da hammadde miktarı yetersizdir. Lütfen ithalat kayıtlarınızı kontrol edin.
                      </p>
                    </div>
                  </div>
                )}
              </div>

              <div className="p-6 border-t border-slate-100 flex justify-between items-center bg-slate-50/50">
                <button 
                  onClick={downloadExportReport}
                  className="flex items-center gap-2 bg-blue-600 text-white px-6 py-3 rounded-xl font-black text-xs uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg active:scale-95"
                >
                  <FileSpreadsheet size={16} />
                  Excel Olarak İndir
                </button>
                <button 
                  onClick={() => setBulkResults(null)}
                  className="bg-slate-900 text-white px-8 py-3 rounded-xl font-black text-xs uppercase tracking-widest hover:bg-slate-800 transition-all shadow-lg"
                >
                  Anladım
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {showImportDialog && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="card bg-white w-full max-w-md p-10 space-y-8"
            >
              <div className="flex items-center gap-4 border-b border-slate-100 pb-4">
                <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-lg flex items-center justify-center">
                  <ArrowDownCircle size={24} />
                </div>
                <h2 className="text-xl font-black text-slate-800 uppercase tracking-tighter">
                  İthalat Kaydı
                </h2>
              </div>
              <div className="space-y-6">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Beyanname No</label>
                    <input 
                      type="text" 
                      value={newImport.declarationNumber}
                      onChange={e => setNewImport({ ...newImport, declarationNumber: e.target.value })}
                      placeholder="Örn: 23063400IM..."
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all font-bold text-sm"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Varış Tarihi</label>
                    <input 
                      type="date" 
                      value={newImport.date}
                      onChange={e => setNewImport({ ...newImport, date: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all font-bold text-sm"
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">GTİP No</label>
                    <input 
                      type="text" 
                      value={newImport.gtip}
                      onChange={e => setNewImport({ ...newImport, gtip: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all font-bold text-sm"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Hammadde Açıklaması</label>
                    <input 
                      type="text" 
                      value={newImport.materialName}
                      onChange={e => setNewImport({ ...newImport, materialName: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all font-bold text-sm"
                    />
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Miktar (KG)</label>
                  <input 
                    type="number" 
                    value={newImport.quantity}
                    onChange={e => setNewImport({ ...newImport, quantity: Number(e.target.value) })}
                    className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl font-black focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all text-lg tracking-tight tabular-nums"
                  />
                </div>
              </div>
              <div className="flex gap-3 pt-4">
                <button 
                  onClick={() => setShowImportDialog(false)}
                  className="flex-1 px-4 py-4 rounded-xl font-black text-xs uppercase tracking-widest text-slate-400 hover:bg-slate-50 transition-colors"
                >
                  Vazgeç
                </button>
                <button 
                  onClick={handleAddImport}
                  className="flex-2 px-4 py-4 bg-blue-600 text-white rounded-xl font-black text-xs uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg shadow-blue-100"
                >
                  Onayla ve Kaydet
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {showExportDialog && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="card bg-white w-full max-w-lg p-10 space-y-8"
            >
              <div className="flex items-center gap-4 border-b border-slate-100 pb-4">
                <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-lg flex items-center justify-center">
                  <ArrowUpCircle size={24} />
                </div>
                <h2 className="text-xl font-black text-slate-800 uppercase tracking-tighter">
                  İhracat & Sarfiyat Eşlemesi
                </h2>
              </div>
              <div className="space-y-6">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">İhracat Beyanname No</label>
                    <input 
                      type="text" 
                      value={newExport.declarationNumber}
                      onChange={e => setNewExport({ ...newExport, declarationNumber: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all font-bold text-sm"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">İhracat/Üretim Tarihi</label>
                    <input 
                      type="date" 
                      value={newExport.date}
                      onChange={e => setNewExport({ ...newExport, date: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all font-bold text-sm"
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Ürün GTİP No</label>
                    <input 
                      type="text" 
                      value={newExport.gtip}
                      onChange={e => setNewExport({ ...newExport, gtip: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all font-bold text-sm"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Satır Kodu (Opsiyonel)</label>
                    <input 
                      type="text" 
                      value={newExport.itemCode}
                      onChange={e => {
                        const code = e.target.value;
                        let rate = newExport.consumptionRate;
                        if (code === '26.1.01378.001') rate = 0.928;
                        else if (code === '26.1.01378.002') rate = 0.865;
                        setNewExport({ ...newExport, itemCode: code, consumptionRate: rate });
                      }}
                      placeholder="Örn: 26.1.01378.001"
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all font-bold text-sm"
                    />
                  </div>
                </div>
                <div className="space-y-2">
                  <div className="flex flex-col gap-2 p-3 bg-blue-50 border border-blue-100 rounded-lg">
                    <div className="flex items-center gap-2">
                      <AlertCircle size={14} className="text-blue-500 shrink-0" />
                      <p className="text-[10px] text-blue-700 font-bold italic leading-tight">
                        Seçilen tarih için kullanılabilir stok: <span className="text-sm not-italic font-black decoration-blue-500/20 underline">{availableStockForSelectedDate.toLocaleString()} KG</span>
                      </p>
                    </div>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Miktar (ADET)</label>
                    <input 
                      type="number" 
                      value={newExport.quantity}
                      onChange={e => setNewExport({ ...newExport, quantity: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl font-black focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all text-lg tabular-nums"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 italic">Sarfiyat Katsayısı</label>
                    <input 
                      type="number" 
                      step="0.01"
                      value={newExport.consumptionRate}
                      onChange={e => setNewExport({ ...newExport, consumptionRate: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-200 p-3 rounded-xl font-black focus:outline-none focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 transition-all text-lg text-blue-600"
                    />
                  </div>
                </div>
                <div className="p-5 bg-slate-800 rounded-2xl flex items-center justify-between border border-white/5 shadow-inner">
                  <div>
                    <p className="text-[10px] font-black text-white/40 uppercase tracking-widest">Hesaplanan İthal Sarfiyatı</p>
                    <p className="text-2xl font-black text-white italic tracking-tighter">{(newExport.quantity * newExport.consumptionRate).toLocaleString()} <span className="text-sm uppercase text-white/40 not-italic">KG</span></p>
                  </div>
                  <div className="w-12 h-12 bg-white/5 rounded-xl flex items-center justify-center text-white/20">
                    <Layers size={24} />
                  </div>
                </div>
              </div>
              <div className="flex gap-3 pt-4">
                <button 
                  onClick={() => setShowExportDialog(false)}
                  className="flex-1 px-4 py-4 rounded-xl font-black text-xs uppercase tracking-widest text-slate-400 hover:bg-slate-50 transition-colors"
                >
                  Vazgeç
                </button>
                <button 
                  onClick={handleAddExport}
                  disabled={availableStockForSelectedDate <= 0}
                  className={`flex-2 px-4 py-4 rounded-xl font-black text-xs uppercase tracking-widest transition-all shadow-lg flex items-center justify-center gap-2 ${
                    availableStockForSelectedDate <= 0 
                    ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none' 
                    : 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-indigo-100'
                  }`}
                >
                  {availableStockForSelectedDate <= 0 ? 'Stok Yetersiz' : 'Sarfiyatı Onayla'}
                  <ArrowRight size={14} />
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function ThemeStatusBadge({ status }: { status: Status }) {
  switch (status) {
    case Status.COMPLETED:
      return (
        <span className="inline-flex items-center gap-1.5 bg-emerald-500/10 text-emerald-500 px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest border border-emerald-500/10">
          <CheckCircle2 size={10} />
          Tamamlandı
        </span>
      );
    case Status.PARTIAL:
      return (
        <span className="inline-flex items-center gap-1.5 bg-amber-500/10 text-amber-500 px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest border border-amber-500/10">
          <Layers size={10} />
          Kısmi
        </span>
      );
    default:
      return (
        <span className="inline-flex items-center gap-1.5 bg-rose-500/10 text-rose-500 px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest border border-rose-500/10">
          <AlertCircle size={10} />
          Bekliyor
        </span>
      );
  }
}
