
import { Import, Export, StockLink, Status } from '../types';

/**
 * DIR Matching Logic
 * 1. Filter imports where import.date < export.date (at least 1 day difference)
 * 2. Use FIFO (sorting by date)
 * 3. Consumed Stock = Export Quantity * Consumption Rate
 */
export const matchExportWithImports = (
  newExport: Export,
  allImports: Import[]
): { updatedImports: Import[]; links: StockLink[]; status: Status } => {
  // Sort imports by date (FIFO)
  const sortedImports = [...allImports].sort(
    (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime()
  );

  const exportDate = new Date(newExport.date);
  let totalToConsume = newExport.quantity * newExport.consumptionRate;
  const originalToConsume = totalToConsume;
  const links: StockLink[] = [];

  const updatedImports = sortedImports.map((imp) => {
    const importDate = new Date(imp.date);
    
    // Check condition: Production cannot happen on the same day as import (at least 1 day diff)
    const timeDiff = exportDate.getTime() - importDate.getTime();
    const daysDiff = timeDiff / (1000 * 3600 * 24);

    if (daysDiff >= 1 && totalToConsume > 0.0001 && imp.remainingQuantity > 0.0001) {
      const canTake = Math.min(imp.remainingQuantity, totalToConsume);
      
      links.push({
        importId: imp.id,
        exportId: newExport.id,
        consumedQuantity: canTake,
        consumptionRate: newExport.consumptionRate,
        exportQuantity: canTake / newExport.consumptionRate
      });

      totalToConsume -= canTake;
      return { ...imp, remainingQuantity: Math.max(0, imp.remainingQuantity - canTake) };
    }
    return imp;
  });

  let status = Status.COMPLETED;
  if (totalToConsume > 0.0001) {
    status = totalToConsume >= originalToConsume - 0.0001 ? Status.PENDING : Status.PARTIAL;
  }

  return { updatedImports, links, status };
};
