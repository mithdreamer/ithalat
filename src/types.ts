
export enum Status {
  PENDING = 'PENDING',
  COMPLETED = 'COMPLETED',
  PARTIAL = 'PARTIAL',
}

export interface Import {
  id: string;
  date: string;
  declarationNumber: string;
  gtip: string;
  materialName: string;
  quantity: number;
  remainingQuantity: number;
  totalCost?: number;
}

export interface Export {
  id: string;
  date: string;
  declarationNumber: string;
  gtip: string;
  itemCode?: string;
  productName: string;
  quantity: number;
  consumptionRate: number; // Consumption rate (e.g., 0.85)
  status: Status;
}

export interface StockLink {
  importId: string;
  exportId: string;
  consumedQuantity: number;
  consumptionRate: number;
  exportQuantity: number;
}
