import React, { useMemo, useState, useCallback, useEffect } from "react";
import { initializeBlock, useColorScheme, useBase, useCustomProperties, useRecords, expandRecord } from "@airtable/blocks/interface/ui";
import { Table, Field } from "@airtable/blocks/interface/models";
import {
  PackageIcon, CubeIcon, CaretRightIcon, PlusIcon, TrashIcon, PencilSimpleIcon,
  XIcon, CheckIcon, WarningIcon, CircleIcon, MagnifyingGlass,
  ShoppingCart, Truck, Calendar, CheckCircle, GearIcon
} from "@phosphor-icons/react";

// --- TYPES ---
type Material = { id: string; name: string; onHand: number; category: string; orderedStr: string; stillToPick: number; };
type Product = { id: string; name: string; };
type Kit = { id: string; name: string; productIds: string[]; items: { productId: string; qty: number }[]; };
// id and moQtyPer added: id lets us match back to bomRecords, moQtyPer is the no-overage qty used only for MO picklist
type BOMEntry = { id: string; productId: string; materialId: string; qtyPer: number; moQtyPer: number; };
type SalesOrderHeader = { id: string; name: string; customer: string; customerPO: string; date: string; status: "Committed" | "TimePhasedDemand"; lineItemIds: string[]; soStatus?: string; rawStatusName?: string; };
type LineItemData = { id: string; productId: string; qty: number; kitId?: string; kitQty?: number; hasAllocation: boolean; isInMO?: boolean; isFullyDone?: boolean; moNames?: string[]; };
type PurchaseOrder = { id: string; poLineId: string; headerId: string; name: string; vendor: string; materialId: string; qty: number; date: string; fallbackDate?: string; status: string; isShipment: boolean; shipmentName?: string; };
type VirtualPO = { id: string; materialId: string; qty: number; date: string; vendor: string; };
type PlanRecord = { id: string; name: string; linkedOrderIds: string[]; status: string; };
type Toast = { id: string; type: "success" | "error"; message: string; };
type MissingConfig = { section: string; items: string[]; };

// --- CUSTOM SESSION HOOK ---
function useSessionState<T>(key: string, initialValue: T): [T, React.Dispatch<React.SetStateAction<T>>] {
  const [state, setState] = useState<T>(() => {
    try {
      const item = window.sessionStorage.getItem(key);
      return item ? JSON.parse(item) : initialValue;
    } catch (error) {
      return initialValue;
    }
  });

  useEffect(() => {
    try {
      window.sessionStorage.setItem(key, JSON.stringify(state));
    } catch (error) {
      console.warn("Session storage restricted", error);
    }
  }, [key, state]);

  return [state, setState];
}

// --- SETTINGS CONFIG ---
function getCustomProperties(base: ReturnType<typeof useBase>) {
  const findT = (str: string) => base.tables.find(t => t.name.toLowerCase().includes(str));
  const materialsTable = findT('material');
  const productsTable = findT('product');
  const bomTable = findT('bom');
  const soLineItemsTable = findT('so line');
  const soHeadersTable = findT('sales order');
  const poHeadersTable = findT('purchase order') ?? findT('po header');
  const poLineItemsTable = findT('po line');
  const poShipmentsTable = base.tables.find(t => t.name.toLowerCase().includes('po shipment'));
  const plansTable = base.tables.find(t => t.name.toLowerCase() === 'planning group') ?? findT('plan');
  const suggestedTable = base.tables.find(t => t.name.toLowerCase() === 'suggested materials' || t.name.toLowerCase() === 'sugested materials') ?? findT('suggest');
  
  const kitsTable = base.tables.find(t => t.name.toLowerCase() === 'kits');
  const kitItemsTable = base.tables.find(t => t.name.toLowerCase() === 'kit items');
  
  const moTable = base.tables.find(t => t.name.toLowerCase().includes('make order'));
  const moAllocTable = base.tables.find(t => t.name.toLowerCase().includes('mo line items'));
  const moPickTable = base.tables.find(t => t.name.toLowerCase().includes('mo picklist') || t.name.toLowerCase().includes('pick list'));

  return [
    { key: 'materialsTable', label: 'Materials Table', type: 'table' as const, defaultValue: materialsTable },
    { key: 'materialsNameField', label: 'Material Name', type: 'field' as const, table: materialsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: materialsTable?.fields.find(f => f.type === 'singleLineText') },
    { key: 'materialsOnHandField', label: 'On Hand Qty', type: 'field' as const, table: materialsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: materialsTable?.fields.find(f => f.name.toLowerCase().includes('qty') || f.name.toLowerCase().includes('available')) },
    { key: 'materialsCategoryField', label: 'Category', type: 'field' as const, table: materialsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: materialsTable?.fields.find(f => f.name.toLowerCase().includes('category')) },
    { key: 'orderedField', label: 'Ordered', type: 'field' as const, table: materialsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: materialsTable?.fields.find(f => f.name.toLowerCase().includes('order')) },
    { key: 'materialsStillToPickField', label: 'Still To Pick (Open MOs)', type: 'field' as const, table: materialsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: materialsTable?.fields.find(f => f.name === 'Still To Pick (Open MOs)') },
    
    { key: 'productsTable', label: 'Products Table', type: 'table' as const, defaultValue: productsTable },
    { key: 'productsNameField', label: 'Product Name', type: 'field' as const, table: productsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: productsTable?.fields.find(f => f.type === 'singleLineText' || f.type === 'multilineText') },
    
    { key: 'kitsTableProp', label: 'Kits Table', type: 'table' as const, defaultValue: kitsTable },
    { key: 'kitsKitItemsField', label: 'Kits -> Kit Items Link', type: 'field' as const, table: kitsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: kitsTable?.fields.find(f => f.name.toLowerCase().includes('kit items')) },
    { key: 'kitItemsTableProp', label: 'Kit Items Table (Junction)', type: 'table' as const, defaultValue: kitItemsTable },
    { key: 'kitItemsProductField', label: 'Kit Item -> Product Link', type: 'field' as const, table: kitItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: kitItemsTable?.fields.find(f => f.name.toLowerCase().includes('product')) },
    { key: 'kitItemsQtyField', label: 'Kit Item Qty Per Kit', type: 'field' as const, table: kitItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: kitItemsTable?.fields.find(f => f.name.toLowerCase().includes('qty')) },

    { key: 'bomTable', label: 'BOM Table', type: 'table' as const, defaultValue: bomTable },
    { key: 'bomProductField', label: 'BOM Product Link', type: 'field' as const, table: bomTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: bomTable?.fields.find(f => f.name.toLowerCase().includes('product')) },
    { key: 'bomMaterialField', label: 'BOM Material Link', type: 'field' as const, table: bomTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: bomTable?.fields.find(f => f.name.toLowerCase().includes('material')) },
    { key: 'bomQtyField', label: 'Materials Needed w/Overage', type: 'field' as const, table: bomTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: bomTable?.fields.find(f => f.name.toLowerCase().includes('needed') || f.name.toLowerCase().includes('qty')) },
    { key: 'bomQtyMOField', label: 'BOM Qty for MO Picklist (No Overage)', type: 'field' as const, table: bomTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: bomTable?.fields.find(f => f.name === 'Materials Needed') },
    
    { key: 'soHeadersTable', label: 'Sales Orders Table', type: 'table' as const, defaultValue: soHeadersTable },
    { key: 'soOrderNameField', label: 'SO Name/Number', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.type === 'singleLineText') },
    { key: 'soOrderCustomerField', label: 'SO Customer', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.name.toLowerCase().includes('customer')) },
    { key: 'soOrderCustomerPOField', label: 'SO Customer PO', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.name.toLowerCase().includes('po') || f.name.toLowerCase().includes('customer po')) },
    { key: 'soOrderDateField', label: 'SO Date', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.name.toLowerCase().includes('date') || f.name.toLowerCase().includes('cancel')) },
    { key: 'soOrderStatusField', label: 'SO Status', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'soTypeField', label: 'SO Type', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.name.toLowerCase().includes('type')) },
    { key: 'soOrderLineItemsField', label: 'SO Line Items Link', type: 'field' as const, table: soHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soHeadersTable?.fields.find(f => f.name.toLowerCase().includes('line')) },
    
    { key: 'soLineItemsTable', label: 'SO Line Items', type: 'table' as const, defaultValue: soLineItemsTable },
    { key: 'soLineItemProductField', label: 'Line Item Product', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('product')) },
    { key: 'soLineItemQtyField', label: 'Line Item Qty', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('qty')) },
    { key: 'soLineItemStatusField', label: 'SOLI Status', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('soli status')) },
    { key: 'soLineItemMoAllocField', label: 'Line Item -> MO Line Items Link', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name === 'MO Line Items') ?? soLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('mo line')) },
    { key: 'soLineItemMoStatusField', label: 'SOLI MO Line Status (Lookup)', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('mo line status') || f.name.toLowerCase().includes('mo status')) },
    { key: 'soLineItemMakeOrdersField', label: 'SOLI -> Make Orders Link', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: soLineItemsTable?.fields.find(f => f.name === 'Make Orders') },
    { key: 'soLineItemKitLinkField', label: 'SOLI -> Kit Link', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'soLineItemKitQtyField', label: 'SOLI Kit Demand Qty', type: 'field' as const, table: soLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },

    { key: 'moTableProp', label: 'Make Orders Table', type: 'table' as const, defaultValue: moTable },
    { key: 'moKitField', label: 'MO -> Kit Link', type: 'field' as const, table: moTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moTable?.fields.find(f => f.name.toLowerCase().includes('kit')) },
    { key: 'moDueDateField', label: 'MO Due Date', type: 'field' as const, table: moTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moTable?.fields.find(f => f.name.toLowerCase().includes('date')) },
    { key: 'moSoHeadersField', label: 'MO -> Sales Orders Link', type: 'field' as const, table: moTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moTable?.fields.find(f => f.name === 'Sales Orders') },
    { key: 'moSoLineItemsField', label: 'MO -> SO Line Items Link', type: 'field' as const, table: moTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moTable?.fields.find(f => f.name === 'SO Line Items') },
    
    { key: 'moAllocTableProp', label: 'MO Allocations Table', type: 'table' as const, defaultValue: moAllocTable },
    { key: 'moAllocMoField', label: 'Alloc -> MO Link', type: 'field' as const, table: moAllocTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moAllocTable?.fields.find(f => f.name.toLowerCase().includes('make order') || f.name.toLowerCase().includes('mo')) },
    { key: 'moAllocSoLineField', label: 'Alloc -> SO Line Item Link', type: 'field' as const, table: moAllocTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moAllocTable?.fields.find(f => f.name.toLowerCase().includes('line')) },
    { key: 'moAllocSoHeaderField', label: 'Alloc -> Sales Order Link', type: 'field' as const, table: moAllocTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moAllocTable?.fields.find(f => f.name === 'Sales Orders') },
    { key: 'moAllocQtyField', label: 'Allocated Qty', type: 'field' as const, table: moAllocTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moAllocTable?.fields.find(f => f.name.toLowerCase().includes('qty') || f.name.toLowerCase().includes('alloc')) },

    { key: 'moPickTableProp', label: 'MO Picklist Table', type: 'table' as const, defaultValue: moPickTable },
    { key: 'moPickMoField', label: 'Pick -> MO Link', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name.toLowerCase().includes('make order') || f.name.toLowerCase().includes('mo')) },
    { key: 'moPickTypeField', label: 'Pick Item Type (Dropdown)', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name.toLowerCase().includes('type')) },
    { key: 'moPickProductField', label: 'Pick -> Product Link', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name.toLowerCase().includes('product')) },
    { key: 'moPickMaterialField', label: 'Pick -> Material Link', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name.toLowerCase().includes('material')) },
    { key: 'moPickQtyField', label: 'Pick Required Qty', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name.toLowerCase().includes('req') || f.name.toLowerCase().includes('qty')) },
    { key: 'moPickSoLineField', label: 'Pick -> SO Line Items Link', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name === 'SO Line Items') },
    { key: 'moPickQtyPerKitField', label: 'Pick Qty Per Kit (Material)', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name === 'Qty Per Kit (Material)') },
    { key: 'moPickStillToPickField', label: 'Pick Still To Pick (Live)', type: 'field' as const, table: moPickTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moPickTable?.fields.find(f => f.name === 'Still To Pick') },
    { key: 'moStatusFieldProp', label: 'MO Status', type: 'field' as const, table: moTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: moTable?.fields.find(f => f.name === 'MO Status') },

    { key: 'poHeadersTable', label: '1. PO Headers Table', type: 'table' as const, defaultValue: poHeadersTable },
    { key: 'poNameField', label: 'PO Name/Number', type: 'field' as const, table: poHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'poVendorField', label: 'PO Vendor', type: 'field' as const, table: poHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'poDateField', label: 'PO Master ETA (Fallback)', type: 'field' as const, table: poHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'poStatusField', label: 'PO Status', type: 'field' as const, table: poHeadersTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    
    { key: 'poLineItemsTable', label: '2. PO Line Items Table', type: 'table' as const, defaultValue: poLineItemsTable },
    { key: 'poHeaderLinkField', label: 'Link to PO Header', type: 'field' as const, table: poLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },
    { key: 'poMaterialField', label: 'PO Material Link', type: 'field' as const, table: poLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('material')) },
    { key: 'poQtyField', label: 'Line Total Qty', type: 'field' as const, table: poLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('qty')) },
    { key: 'poLineItemRemainingQtyField', label: 'Line Qty Unscheduled', type: 'field' as const, table: poLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poLineItemsTable?.fields.find(f => f.name.toLowerCase().includes('unsched')) },
    { key: 'poLineItemDateField', label: 'Line ETA Date (Optional)', type: 'field' as const, table: poLineItemsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true },

    { key: 'poShipmentsTable', label: '3. PO Shipments Table', type: 'table' as const, defaultValue: poShipmentsTable },
    { key: 'poShipmentLinkField', label: 'Link to PO Line', type: 'field' as const, table: poShipmentsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poShipmentsTable?.fields.find(f => f.name.toLowerCase().includes('line')) },
    { key: 'poShipmentQtyField', label: 'Shipment Qty (e.g. En Rout)', type: 'field' as const, table: poShipmentsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poShipmentsTable?.fields.find(f => f.name.toLowerCase().includes('qty')) },
    { key: 'poShipmentDateField', label: 'Shipment ETA Date', type: 'field' as const, table: poShipmentsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poShipmentsTable?.fields.find(f => f.name.toLowerCase().includes('date') || f.name.toLowerCase().includes('eta')) },
    { key: 'poShipmentStatusField', label: 'Shipment Status (Kill Switch)', type: 'field' as const, table: poShipmentsTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: poShipmentsTable?.fields.find(f => f.name.toLowerCase().includes('status')) },

    { key: 'plansTable', label: 'Plans Table', type: 'table' as const, defaultValue: plansTable },
    { key: 'planNameField', label: 'Plan Name', type: 'field' as const, table: plansTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: plansTable?.fields.find(f => f.type === 'singleLineText') },
    { key: 'planOrdersLinkField', label: 'Plan Orders Link', type: 'field' as const, table: plansTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: plansTable?.fields.find(f => f.name.toLowerCase().includes('sales') || f.name.toLowerCase().includes('order')) },
    { key: 'planStatusField', label: 'Plan Status', type: 'field' as const, table: plansTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: plansTable?.fields.find(f => f.name.toLowerCase().includes('status')) },
    
    { key: 'planVpoTable', label: 'Plan Virtual POs', type: 'table' as const, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual')) },
    { key: 'planVpoPlanField', label: 'VPO Plan Link', type: 'field' as const, table: base.tables.find(t => t.name.toLowerCase().includes('virtual')) ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual'))?.fields.find(f => f.name.toLowerCase().includes('plan')) },
    { key: 'planVpoMaterialField', label: 'VPO Material Link', type: 'field' as const, table: base.tables.find(t => t.name.toLowerCase().includes('virtual')) ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual'))?.fields.find(f => f.name.toLowerCase().includes('material')) },
    { key: 'planVpoQtyField', label: 'VPO Qty', type: 'field' as const, table: base.tables.find(t => t.name.toLowerCase().includes('virtual')) ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual'))?.fields.find(f => f.name.toLowerCase().includes('qty')) },
    { key: 'planVpoDateField', label: 'VPO Date', type: 'field' as const, table: base.tables.find(t => t.name.toLowerCase().includes('virtual')) ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual'))?.fields.find(f => f.name.toLowerCase().includes('date')) },
    { key: 'planVpoVendorField', label: 'VPO Vendor', type: 'field' as const, table: base.tables.find(t => t.name.toLowerCase().includes('virtual')) ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: base.tables.find(t => t.name.toLowerCase().includes('virtual'))?.fields.find(f => f.name.toLowerCase().includes('vendor')) },
    
    { key: 'suggestedTable', label: 'Suggested Materials Table', type: 'table' as const, defaultValue: suggestedTable },
    { key: 'suggestedPlanField', label: 'Planning Group Link Field', type: 'field' as const, table: suggestedTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: suggestedTable?.fields.find(f => f.name === 'Planning Group') },
    { key: 'suggestedMaterialField', label: 'Materials Link Field', type: 'field' as const, table: suggestedTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: suggestedTable?.fields.find(f => f.name === 'Materials') },
    { key: 'suggestedQtyField', label: 'Qty Suggested Field', type: 'field' as const, table: suggestedTable ?? base.tables[0]!, shouldFieldBeAllowed: () => true, defaultValue: suggestedTable?.fields.find(f => f.name.includes('Qty Suggested') || f.name.toLowerCase().includes('qty')) },
  ];
}

// --- UTILS ---
function formatDate(d: string) {
  if (!d) return "";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt.getTime())) return "";
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function uid() { 
  return typeof crypto !== 'undefined' && crypto.randomUUID 
    ? crypto.randomUUID() 
    : `id-${Date.now()}-${Math.floor(Math.random() * 1e9)}`; 
}

function parseStrictDate(dateValue: unknown): string {
  if (!dateValue) return "";
  let val: unknown = Array.isArray(dateValue) ? dateValue[0] : dateValue;
  if (val && typeof val === "object" && "date" in val) {
    val = (val as { date?: string }).date;
  }
  if (typeof val === "string") {
    const match = /^(\d{4}-\d{2}-\d{2})/.exec(val);
    if (match?.[1]) {
      const testDate = new Date(match[1] + "T00:00:00");
      if (!isNaN(testDate.getTime())) return match[1];
    }
  }
  return "";
}

const getVal = (r: any, f?: Field) => f?.id ? r.getCellValue(f.id) : null;
const getStr = (r: any, f?: Field) => f?.id ? (r.getCellValueAsString(f.id) || r.name) : r.name;
const getLinks = (r: any, f?: Field) => (getVal(r, f) as any[]) || [];
const getFirstLinkId = (r: any, f?: Field) => getLinks(r, f)[0]?.id ?? '';

function ConfigurationError({ missingConfigs }: { missingConfigs: MissingConfig[] }) {
  const { colorScheme } = useColorScheme();
  const isDark = colorScheme === "dark";
  const bgMain = isDark ? "bg-gray-800" : "bg-gray-50";
  const bgCard = isDark ? "bg-gray-700" : "bg-white";
  const borderColor = isDark ? "border-gray-600" : "border-gray-200";
  const textPrimary = isDark ? "text-gray-100" : "text-gray-900";
  const textSecondary = isDark ? "text-gray-400" : "text-gray-500";
  const configBg = isDark ? "bg-gray-600" : "bg-gray-50";

  return (
    <main className={`min-h-screen w-full flex items-center justify-center ${bgMain} ${textPrimary} p-8`}>
      <div className={`${bgCard} rounded-lg border ${borderColor} shadow-lg max-w-2xl w-full p-6`}>
        <div className="flex items-center gap-3 mb-6">
          <div className="p-3 rounded-full bg-rose-100 dark:bg-rose-900">
            <WarningIcon weight="bold" className="w-6 h-6 text-rose-600 dark:text-rose-400" />
          </div>
          <div>
            <h1 className="text-xl font-bold">Configuration Required</h1>
            <p className={`text-sm ${textSecondary}`}>Please configure the missing fields in the properties panel (gear icon).</p>
          </div>
        </div>
        <div className="space-y-4">
          {missingConfigs.map((config) => (
            <div key={config.section} className={`p-4 rounded-lg border ${borderColor} ${configBg}`}>
              <h3 className="font-semibold text-sm mb-2">{config.section}</h3>
              <ul className="space-y-1">
                {config.items.map((item) => (
                  <li key={item} className={`text-sm ${textSecondary} flex items-center gap-2`}>
                    <XIcon weight="bold" className="w-3.5 h-3.5 text-rose-500" /> {item}
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </div>
    </main>
  );
}

function SupplyChainPlanner(): React.ReactElement {
  const base = useBase();
  const { colorScheme } = useColorScheme();
  const isDark = colorScheme === "dark";
  
  const getPropsConfig = useCallback(() => getCustomProperties(base), [base]);
  const { customPropertyValueByKey: props } = useCustomProperties(getPropsConfig);

  // TABLES & FIELDS
  const materialsTable = props.materialsTable as Table | undefined;
  const materialsNameField = props.materialsNameField as Field | undefined;
  const materialsOnHandField = props.materialsOnHandField as Field | undefined;
  const materialsCategoryField = props.materialsCategoryField as Field | undefined;
  const orderedField = props.orderedField as Field | undefined;
  const materialsStillToPickField = props.materialsStillToPickField as Field | undefined;
  
  const productsTable = props.productsTable as Table | undefined;
  const productsNameField = props.productsNameField as Field | undefined;

  const kitsTableProp = props.kitsTableProp as Table | undefined;
  const kitsKitItemsField = props.kitsKitItemsField as Field | undefined;
  const kitItemsTableProp = props.kitItemsTableProp as Table | undefined;
  const kitItemsProductField = props.kitItemsProductField as Field | undefined;
  const kitItemsQtyField = props.kitItemsQtyField as Field | undefined;
  
  const bomTable = props.bomTable as Table | undefined;
  const bomProductField = props.bomProductField as Field | undefined;
  const bomMaterialField = props.bomMaterialField as Field | undefined;
  const bomQtyField = props.bomQtyField as Field | undefined;
  const bomQtyMOField = props.bomQtyMOField as Field | undefined;

  const plansTable = props.plansTable as Table | undefined;
  const planNameField = props.planNameField as Field | undefined;
  const planOrdersLinkField = props.planOrdersLinkField as Field | undefined;
  const planStatusField = props.planStatusField as Field | undefined;
  
  const soHeadersTable = props.soHeadersTable as Table | undefined;
  const soHeadersView = useMemo(() => { try { return soHeadersTable?.getViewByName("Open Orders") || null; } catch { return null; } }, [soHeadersTable]);
  const soOrderNameField = props.soOrderNameField as Field | undefined;
  const soOrderCustomerField = props.soOrderCustomerField as Field | undefined;
  const soOrderCustomerPOField = props.soOrderCustomerPOField as Field | undefined;
  const soOrderDateField = props.soOrderDateField as Field | undefined;
  const soOrderStatusField = props.soOrderStatusField as Field | undefined;
  const soOrderLineItemsField = props.soOrderLineItemsField as Field | undefined;
  const soTypeField = props.soTypeField as Field | undefined;
  
  const soLineItemsTable = props.soLineItemsTable as Table | undefined;
  const soLineItemProductField = props.soLineItemProductField as Field | undefined;
  const soLineItemQtyField = props.soLineItemQtyField as Field | undefined;
  const soLineItemStatusField = props.soLineItemStatusField as Field | undefined;
  const soLineItemMoAllocField = props.soLineItemMoAllocField as Field | undefined;
  const soLineItemMoStatusField = props.soLineItemMoStatusField as Field | undefined;
  const soLineItemMakeOrdersField = props.soLineItemMakeOrdersField as Field | undefined;
  const soLineItemKitLinkField = props.soLineItemKitLinkField as Field | undefined;
  const soLineItemKitQtyField = props.soLineItemKitQtyField as Field | undefined;

  const moTableProp = props.moTableProp as Table | undefined;
  const moKitField = props.moKitField as Field | undefined;
  const moDueDateField = props.moDueDateField as Field | undefined;
  const moSoHeadersField = props.moSoHeadersField as Field | undefined;
  const moSoLineItemsField = props.moSoLineItemsField as Field | undefined;

  const moAllocTableProp = props.moAllocTableProp as Table | undefined;
  const moAllocMoField = props.moAllocMoField as Field | undefined;
  const moAllocSoLineField = props.moAllocSoLineField as Field | undefined;
  const moAllocSoHeaderField = props.moAllocSoHeaderField as Field | undefined;
  const moAllocQtyField = props.moAllocQtyField as Field | undefined;

  const moPickTableProp = props.moPickTableProp as Table | undefined;
  const moPickMoField = props.moPickMoField as Field | undefined;
  const moPickTypeField = props.moPickTypeField as Field | undefined;
  const moPickProductField = props.moPickProductField as Field | undefined;
  const moPickMaterialField = props.moPickMaterialField as Field | undefined;
  const moPickQtyField = props.moPickQtyField as Field | undefined;
  const moPickSoLineField = props.moPickSoLineField as Field | undefined;
  const moPickQtyPerKitField = props.moPickQtyPerKitField as Field | undefined;
  const moPickStillToPickField = props.moPickStillToPickField as Field | undefined;
  const moStatusFieldProp = props.moStatusFieldProp as Field | undefined;
  
  const poHeadersTable = props.poHeadersTable as Table | undefined;
  const poHeadersView = useMemo(() => { try { return poHeadersTable?.getViewByName("Open POs") || null; } catch { return null; } }, [poHeadersTable]);
  const poNameField = props.poNameField as Field | undefined;
  const poVendorField = props.poVendorField as Field | undefined;
  const poDateField = props.poDateField as Field | undefined;
  const poStatusField = props.poStatusField as Field | undefined;

  const poLineItemsTable = props.poLineItemsTable as Table | undefined;
  const poLineItemsView = useMemo(() => { try { return poLineItemsTable?.getViewByName("Open Orders") || null; } catch { return null; } }, [poLineItemsTable]);
  const poHeaderLinkField = props.poHeaderLinkField as Field | undefined;
  const poMaterialField = props.poMaterialField as Field | undefined;
  const poQtyField = props.poQtyField as Field | undefined;
  const poLineItemRemainingQtyField = props.poLineItemRemainingQtyField as Field | undefined;
  const poLineItemDateField = props.poLineItemDateField as Field | undefined;

  const poShipmentsTable = props.poShipmentsTable as Table | undefined;
  const poShipmentLinkField = props.poShipmentLinkField as Field | undefined;
  const poShipmentQtyField = props.poShipmentQtyField as Field | undefined;
  const poShipmentDateField = props.poShipmentDateField as Field | undefined;
  const poShipmentStatusField = props.poShipmentStatusField as Field | undefined;
  
  const planVpoTable = props.planVpoTable as Table | undefined;
  const planVpoPlanField = props.planVpoPlanField as Field | undefined;
  const planVpoMaterialField = props.planVpoMaterialField as Field | undefined;
  const planVpoQtyField = props.planVpoQtyField as Field | undefined;
  const planVpoDateField = props.planVpoDateField as Field | undefined;
  const planVpoVendorField = props.planVpoVendorField as Field | undefined;

  const suggestedTable = props.suggestedTable as Table | undefined;
  const suggestedPlanField = props.suggestedPlanField as Field | undefined;
  const suggestedMaterialField = props.suggestedMaterialField as Field | undefined;
  const suggestedQtyField = props.suggestedQtyField as Field | undefined;

  const missingConfigs = useMemo(() => {
    const missing: MissingConfig[] = [];
    if (!materialsTable || !materialsNameField?.id || !materialsOnHandField?.id) {
      missing.push({ section: "Materials", items: ["Table, Name, or On Hand Qty missing"] });
    }
    if (!productsTable || !productsNameField?.id) {
      missing.push({ section: "Products", items: ["Table or Name field missing"] });
    }
    if (kitsTableProp && (!kitsKitItemsField?.id || !kitItemsTableProp || !kitItemsProductField?.id)) {
      missing.push({ section: "Kits", items: ["Please map the Kit Items junction table and fields in settings"] });
    }
    if (!soLineItemsTable || !soLineItemMoAllocField?.id || !soLineItemKitLinkField?.id || !soLineItemKitQtyField?.id) {
  missing.push({ section: "Make Order Linking", items: ["Map 'MO Allocations Link', 'Kit Link', and 'Kit Demand Qty' in settings"] });
  }
    if (!moTableProp || !moKitField?.id || !moAllocTableProp || !moAllocMoField?.id || !moPickTableProp || !moPickQtyField?.id) {
      missing.push({ section: "Make Orders Base", items: ["Please map Make Orders, Allocations, and Picklist tables in settings"] });
    }
    if (!bomTable || !bomProductField?.id || !bomMaterialField?.id || !bomQtyField?.id) {
      missing.push({ section: "BOM", items: ["Table or Link/Qty fields missing"] });
    }
    if (!soHeadersTable || !soOrderNameField?.id || !soOrderDateField?.id) {
      missing.push({ section: "Sales Orders", items: ["Table, Name, or Date missing"] });
    }
    if (!soLineItemsTable || !soLineItemProductField?.id || !soLineItemQtyField?.id) {
      missing.push({ section: "SO Line Items", items: ["Table, Product Link, or Qty missing"] });
    }
    return missing;
  }, [
    materialsTable, materialsNameField, materialsOnHandField,
    productsTable, productsNameField, kitsTableProp, kitsKitItemsField, kitItemsTableProp, kitItemsProductField,
    soLineItemsTable, soLineItemMoAllocField, soLineItemKitLinkField, soLineItemKitQtyField, moTableProp, moKitField, moAllocTableProp, moAllocMoField, moPickTableProp, moPickQtyField,
    bomTable, bomProductField, bomMaterialField, bomQtyField,
    soHeadersTable, soOrderNameField, soOrderDateField,
    soLineItemsTable, soLineItemProductField, soLineItemQtyField
  ]);

  // DATA FETCHING
  const materialsRecords = useRecords(materialsTable ?? null) ?? [];
  const productsRecords = useRecords(productsTable ?? null) ?? [];
  const bomRecords = useRecords(bomTable ?? null) ?? [];
  const soHeaderRecords = useRecords((soHeadersView || soHeadersTable) ?? null) ?? [];
  const soLineItemRecords = useRecords(soLineItemsTable ?? null) ?? [];
  const moRecords = useRecords(moTableProp ?? null) ?? [];
  const moPickRecords = useRecords(moPickTableProp ?? null) ?? [];
  const poHeaderRecords = useRecords((poHeadersView || poHeadersTable) ?? null) ?? [];
  const poLineItemRecords = useRecords((poLineItemsView || poLineItemsTable) ?? null) ?? [];
  const poShipmentRecords = useRecords(poShipmentsTable ?? null) ?? [];
  const planRecords = useRecords(plansTable ?? null) ?? [];
  const planVpoRecords = useRecords(planVpoTable ?? null) ?? [];
  const suggestedRecords = useRecords(suggestedTable ?? null) ?? [];
  
  const kitsRecords = useRecords(kitsTableProp ?? null) ?? [];
  const kitItemsRecords = useRecords(kitItemsTableProp ?? null) ?? [];

  // MAPPING
  const materials = useMemo((): Material[] => {
    return materialsRecords.map((r) => {
      let onHand = 0; const rawHand = getVal(r, materialsOnHandField);
      if (typeof rawHand === 'number') onHand = rawHand;
      if (Array.isArray(rawHand) && typeof rawHand[0] === 'number') onHand = rawHand[0];
      const cat = getVal(r, materialsCategoryField);
      const orderedStr = orderedField?.id ? (r.getCellValueAsString(orderedField.id) || '') : '';
      let stillToPick = 0; const rawStp = getVal(r, materialsStillToPickField);
      if (typeof rawStp === 'number') stillToPick = rawStp;
      if (Array.isArray(rawStp) && typeof rawStp[0] === 'number') stillToPick = rawStp[0];
      return { id: r.id, name: getStr(r, materialsNameField), onHand, category: cat?.name ?? 'Uncategorized', orderedStr, stillToPick };
    });
  }, [materialsRecords, materialsNameField, materialsOnHandField, materialsCategoryField, orderedField, materialsStillToPickField]);

  const products = useMemo((): Product[] => {
    return productsRecords.map((r) => ({ id: r.id, name: getStr(r, productsNameField) }));
  }, [productsRecords, productsNameField]);

  // THE TRUE KIT FIX: Mapping through the Kit Items Junction Table
  const kits = useMemo((): Kit[] => {
    if (!kitsTableProp || !kitsKitItemsField || !kitItemsTableProp || !kitItemsProductField) return [];
    
    // NEW: Searches specifically for your 'Kit Name' column
    const kitNameField = kitsTableProp.getFieldIfExists('Kit Name');

    return kitsRecords.map((r) => {
      // NEW: Uses 'Kit Name' if found, otherwise falls back to the default column
      const explicitName = kitNameField ? (r.getCellValueAsString(kitNameField.id) || '').trim() : '';
      const name = explicitName || r.name || 'Unnamed Kit';
      
      const kitItemLinks = getLinks(r, kitsKitItemsField);
      const productIdsSet = new Set<string>();
      const items: { productId: string, qty: number }[] = [];

      kitItemLinks.forEach(link => {
          const kiRecord = kitItemsRecords.find(ki => ki.id === link.id);
          if (kiRecord) {
              const productId = getFirstLinkId(kiRecord, kitItemsProductField);
              if (productId) {
                  productIdsSet.add(productId);
                  let qty = 1;
                  if (kitItemsQtyField) {
                      const rawQty = getVal(kiRecord, kitItemsQtyField);
                      if (typeof rawQty === 'number') qty = rawQty;
                  }
                  items.push({ productId, qty });
              }
          }
      });

      return { id: r.id, name, productIds: Array.from(productIdsSet), items };
    });
  }, [kitsTableProp, kitsRecords, kitsKitItemsField, kitItemsTableProp, kitItemsRecords, kitItemsProductField, kitItemsQtyField]);

  const soLineItemKitData = useMemo(() => {
    const map = new Map<string, { kitId: string; kitQty: number }>();
    
    // If the table or the link field isn't mapped, bail out early
    if (!soLineItemsTable || !soLineItemKitLinkField) return map;

    soLineItemRecords.forEach((r) => {
      // 1. Read dynamic Link Field
      const kitLinks = (r.getCellValue(soLineItemKitLinkField.id) as any[]) || [];
      if (kitLinks.length === 0) return;
      const kitId = kitLinks[0].id;

      // 2. Read dynamic Qty Field
      let kitQty = 0;
      if (soLineItemKitQtyField) {
        const raw = r.getCellValue(soLineItemKitQtyField.id);
        if (typeof raw === 'number') kitQty = raw;
        if (Array.isArray(raw) && raw.length > 0 && typeof raw[0] === 'number') kitQty = raw[0];
      }
      
      if (kitQty > 0) map.set(r.id, { kitId, kitQty });
    });
    return map;
  }, [soLineItemsTable, soLineItemRecords, soLineItemKitLinkField, soLineItemKitQtyField]);

  // KEY CHANGE: bom now includes id (for reference) and moQtyPer (no-overage qty for MO picklist).
  // qtyPer = Materials Needed w/Overage -> used for waterfall / demand planning (unchanged)
  // moQtyPer = Materials Needed (no overage) -> used ONLY when creating MO picklist records
  const bom = useMemo((): BOMEntry[] => {
    return bomRecords.map((r) => {
      let qty = 0;
      const rawQty = getVal(r, bomQtyField);
      if (typeof rawQty === 'number') qty = rawQty;
      if (Array.isArray(rawQty) && typeof rawQty[0] === 'number') qty = rawQty[0];

      // Read the no-overage field if configured, otherwise fall back to same qty
      let moQty = qty;
      if (bomQtyMOField?.id) {
        const raw = r.getCellValue(bomQtyMOField.id);
        if (typeof raw === 'number') moQty = raw;
      }

      return {
        id: r.id,
        productId: getFirstLinkId(r, bomProductField),
        materialId: getFirstLinkId(r, bomMaterialField),
        qtyPer: qty,       // with overage — used for planning/waterfall
        moQtyPer: moQty,   // without overage — used only for MO picklist
      };
    }).filter(b => b.productId && b.materialId);
  }, [bomRecords, bomProductField, bomMaterialField, bomQtyField, bomQtyMOField]);

  const lineItemDataById = useMemo(() => {
    const map = new Map<string, LineItemData>();
    soLineItemRecords.forEach((r) => {
      const rawStatus = soLineItemStatusField ? (r.getCellValueAsString(soLineItemStatusField.id) || "") : "";
      const statusKey = rawStatus.trim().toLowerCase();
      // Hide only truly gone lines; "Released MO" stays visible (shown in MO color with its MO number)
      if (statusKey === 'shipped' || statusKey === 'canceld' || statusKey === 'canceled' || statusKey === 'cancelled') return;
      const isReleasedStatus = statusKey === 'released mo';

      const rawAlloc = soLineItemMoAllocField ? r.getCellValue(soLineItemMoAllocField.id) : null;
      const allocLinks: any[] = Array.isArray(rawAlloc) ? rawAlloc : [];
      // Released lines count as allocated even if the allocation-link field isn't mapped
      const hasAllocation = allocLinks.length > 0 || isReleasedStatus;
      // Prefer real MO numbers (SOLI -> Make Orders link); fall back to MOLI names
      const moLinkRaw = soLineItemMakeOrdersField ? r.getCellValue(soLineItemMakeOrdersField.id) : null;
      const moLinks: any[] = Array.isArray(moLinkRaw) ? moLinkRaw : [];
      const moNames = (moLinks.length > 0 ? moLinks : allocLinks).map((l: any) => l?.name).filter(Boolean);

      // "In MO" until ALL linked MO Line Statuses read "Done"; then it hides.
      let isFullyDone: boolean;
      if (soLineItemMoStatusField?.id) {
        const statusStr = r.getCellValueAsString(soLineItemMoStatusField.id) || '';
        const statuses = statusStr.split(',').map(s => s.trim()).filter(Boolean);
        const allDone = statuses.length > 0 && statuses.every(s => s.toLowerCase() === 'done');
        isFullyDone = hasAllocation && allDone;
      } else {
        // No status field mapped → fall back to legacy behavior (hide once allocated)
        isFullyDone = hasAllocation;
      }
      const isInMO = hasAllocation && !isFullyDone;

      let qty = 0; const rawQty = getVal(r, soLineItemQtyField);
      if (typeof rawQty === 'number') qty = rawQty;
      if (Array.isArray(rawQty) && typeof rawQty[0] === 'number') qty = rawQty[0];
      
      const pId = getFirstLinkId(r, soLineItemProductField);
      const kitData = soLineItemKitData.get(r.id);

      if (pId) {
        map.set(r.id, { id: r.id, productId: pId, qty, kitId: kitData?.kitId, kitQty: kitData?.kitQty, hasAllocation, isInMO, isFullyDone, moNames });
      } else if (kitData?.kitId && (kitData.kitQty ?? 0) > 0) {
        map.set(r.id, { id: r.id, productId: '', qty: 0, kitId: kitData.kitId, kitQty: kitData.kitQty, hasAllocation, isInMO, isFullyDone, moNames });
      }
    });
    return map;
  }, [soLineItemsTable, soLineItemRecords, soLineItemProductField, soLineItemQtyField, soLineItemStatusField, soLineItemKitData, soLineItemMoAllocField, soLineItemMoStatusField, soLineItemMakeOrdersField]);

  const salesOrderHeaders = useMemo((): SalesOrderHeader[] => {
    return soHeaderRecords.map((r) => {
      const name = getStr(r, soOrderNameField);
      let customer = name;
      const cVal = getVal(r, soOrderCustomerField);
      if (typeof cVal === 'string') customer = cVal;
      else if (Array.isArray(cVal) && cVal.length > 0) customer = cVal[0]?.name ?? name;
      
      const customerPO = soOrderCustomerPOField?.id ? (r.getCellValueAsString(soOrderCustomerPOField.id) || '') : '';
      
      let status: "Committed" | "TimePhasedDemand" = "TimePhasedDemand";
      let rawStatus = '';
      const sVal = getVal(r, soOrderStatusField);
      if (typeof sVal === 'string') rawStatus = sVal;
      else if (sVal && typeof sVal === 'object' && 'name' in sVal) rawStatus = sVal.name;
      if (rawStatus.trim().toLowerCase().includes('commit')) status = 'Committed';
      
      const tStat = rawStatus.trim().toLowerCase();
      if (tStat.includes('ship') || tStat.includes('cancel') || tStat.includes('delete')) return null;
      
      return { id: r.id, name, customer, customerPO, date: parseStrictDate(getVal(r, soOrderDateField)), status, lineItemIds: getLinks(r, soOrderLineItemsField).map(l=>l.id), soStatus: rawStatus || 'Pending', rawStatusName: rawStatus };
    }).filter((so): so is SalesOrderHeader => so !== null && !!so.date);
  }, [soHeaderRecords, soOrderNameField, soOrderDateField, soOrderCustomerField, soOrderCustomerPOField, soOrderStatusField, soOrderLineItemsField]);

  const plans = useMemo((): PlanRecord[] => {
    return planRecords.map((r) => {
      let status = 'Draft';
      const sVal = getVal(r, planStatusField);
      if (sVal?.name) status = sVal.name;
      return { id: r.id, name: getStr(r, planNameField), linkedOrderIds: getLinks(r, planOrdersLinkField).map(l=>l.id), status };
    });
  }, [planRecords, planNameField, planOrdersLinkField, planStatusField]);

  const soTypeIndex = useMemo(() => {
    const map = new Map<string, string[]>();
    const display = new Map<string, string>();
    const field = soTypeField ?? soHeadersTable?.getFieldIfExists?.('SO Type') ?? undefined;
    if (!field) return { values: [] as { key: string; label: string }[], map, found: false };
    const recById = new Map(soHeaderRecords.map(r => [r.id, r] as [string, typeof r]));
    salesOrderHeaders.forEach(so => {
      const rec = recById.get(so.id);
      if (!rec) return;
      const str = rec.getCellValueAsString(field.id) || '';
      str.split(',').map(p => p.trim()).filter(Boolean).forEach(name => {
        const k = name.toLowerCase();
        if (!display.has(k)) display.set(k, name);
        const arr = map.get(k) ?? [];
        arr.push(so.id);
        map.set(k, arr);
      });
    });
    const values = Array.from(display.entries()).map(([key, label]) => ({ key, label })).sort((a, b) => a.label.localeCompare(b.label));
    return { values, map, found: true };
  }, [salesOrderHeaders, soHeaderRecords, soTypeField, soHeadersTable]);

  const purchaseOrders = useMemo((): PurchaseOrder[] => {
    if (!poLineItemsTable || !poHeaderLinkField?.id || !poMaterialField?.id || !poQtyField?.id) return [];
    const headerMap = new Map();
    poHeaderRecords.forEach(h => headerMap.set(h.id, h));

    const shipmentsByLine = new Map<string, any[]>();
    if (poShipmentLinkField?.id) {
        poShipmentRecords.forEach(s => {
            const linkId = getFirstLinkId(s, poShipmentLinkField);
            if (linkId) {
                if (!shipmentsByLine.has(linkId)) shipmentsByLine.set(linkId, []);
                shipmentsByLine.get(linkId)!.push(s);
            }
        });
    }

    const flatPOs: PurchaseOrder[] = [];

    poLineItemRecords.forEach((r) => {
      const headerLinks = getLinks(r, poHeaderLinkField);
      const headerRecord = headerLinks.length > 0 ? headerMap.get(headerLinks[0].id) : null;
      if (!headerRecord) return;

      let poName = getStr(headerRecord, poNameField);
      let vendor = 'Unknown Vendor';
      const vVal = getVal(headerRecord, poVendorField);
      if (typeof vVal === 'string') vendor = vVal;
      else if (Array.isArray(vVal) && vVal.length > 0) vendor = vVal[0]?.name ?? 'Unknown Vendor';
      
      let status = 'Ordered';
      const sVal = getVal(headerRecord, poStatusField);
      if (typeof sVal === 'string') status = sVal;
      else if (sVal?.name) status = sVal.name;

      if (status.toLowerCase().includes('received') || status.toLowerCase().includes('closed') || status.toLowerCase().includes('completed')) return; 

      const materialId = getFirstLinkId(r, poMaterialField);
      if (!materialId) return;

      let totalQty = 0;
      const rawTotalQty = getVal(r, poQtyField);
      if (typeof rawTotalQty === 'number') totalQty = rawTotalQty;
      if (Array.isArray(rawTotalQty) && rawTotalQty.length > 0 && typeof rawTotalQty[0] === 'number') totalQty = rawTotalQty[0];

      const shipments = shipmentsByLine.get(r.id) || [];
      let shippedSoFar = 0; 
      const headerDate = poDateField?.id ? parseStrictDate(getVal(headerRecord, poDateField)) : '';

      shipments.forEach(shipment => {
          let isKilled = false;
          if (poShipmentStatusField?.id) {
              const rawStat = getVal(shipment, poShipmentStatusField);
              let statStr = '';
              if (typeof rawStat === 'string') statStr = rawStat.toLowerCase();
              else if (rawStat && typeof rawStat === 'object' && 'name' in rawStat) statStr = rawStat.name.toLowerCase();
              if (statStr.includes('received') || statStr.includes('completed') || statStr.includes('closed') || statStr.includes('cancel')) isKilled = true;
          }

          let sQty = 0;
          const rawSQty = poShipmentQtyField?.id ? getVal(shipment, poShipmentQtyField) : null;
          if (typeof rawSQty === 'number') sQty = rawSQty;
          if (Array.isArray(rawSQty) && rawSQty.length > 0 && typeof rawSQty[0] === 'number') sQty = rawSQty[0];
          
          shippedSoFar += sQty;
          if (isKilled) sQty = 0;
          
          const sDate = poShipmentDateField?.id ? parseStrictDate(getVal(shipment, poShipmentDateField)) : '';
          const shipmentName = getStr(shipment);

          if (sQty > 0 || shipments.length > 0) {
              flatPOs.push({ id: shipment.id, poLineId: r.id, headerId: headerRecord.id, name: poName, vendor, materialId, qty: sQty, date: sDate, fallbackDate: headerDate, status, isShipment: true, shipmentName });
          }
      });

      let remainingQty = 0;
      if (poLineItemRemainingQtyField?.id) {
          const rawRem = getVal(r, poLineItemRemainingQtyField);
          if (typeof rawRem === 'number') remainingQty = rawRem;
          if (Array.isArray(rawRem) && rawRem.length > 0 && typeof rawRem[0] === 'number') remainingQty = rawRem[0];
      } else {
          remainingQty = totalQty - shippedSoFar;
      }

      const lineDate = poLineItemDateField?.id ? parseStrictDate(getVal(r, poLineItemDateField)) : '';
      const fallbackForUnscheduled = lineDate || headerDate;
        
      if (remainingQty > 0 || shipments.length === 0) {
          flatPOs.push({ id: r.id, poLineId: r.id, headerId: headerRecord.id, name: poName, vendor, materialId, qty: remainingQty, date: '', fallbackDate: fallbackForUnscheduled, status, isShipment: false });
      }
    });

    return flatPOs;
  }, [poLineItemsTable, poLineItemRecords, poHeaderRecords, poShipmentRecords, poHeaderLinkField, poNameField, poMaterialField, poQtyField, poLineItemDateField, poDateField, poVendorField, poStatusField, poShipmentLinkField, poShipmentQtyField, poShipmentDateField, poShipmentStatusField, poLineItemRemainingQtyField]);

  const suggestedBuys = useMemo(() => {
    return suggestedRecords.map(r => {
      let qty = 0; const rawQty = getVal(r, suggestedQtyField);
      if (typeof rawQty === 'number') qty = rawQty;
      if (Array.isArray(rawQty) && typeof rawQty[0] === 'number') qty = rawQty[0];
      return { id: r.id, materialId: getFirstLinkId(r, suggestedMaterialField), planId: getFirstLinkId(r, suggestedPlanField) || null, qty };
    });
  }, [suggestedRecords, suggestedMaterialField, suggestedQtyField, suggestedPlanField]);


  // STATE MANAGEMENT
  const [selectedPlanId, setSelectedPlanId] = useSessionState<string>("sc_planId", "");
  const [selectedSOIds, setSelectedSOIds] = useSessionState<string[]>("sc_soIds", []);
  const [dateCutoff, setDateCutoff] = useSessionState<string>("sc_dateCutoff", "");
  
  const [showFG, setShowFG] = useState<boolean>(false);
  const [showRM, setShowRM] = useState<boolean>(false);
  const [showKits, setShowKits] = useState<boolean>(false); 

  const [categoryFilter, setCategoryFilter] = useSessionState<string[]>("sc_catFilter", []);
  const [strictMaterialFilter, setStrictMaterialFilter] = useSessionState<boolean>("sc_strictMat", false);
  const [materialFilter, setMaterialFilter] = useSessionState<string[]>("sc_matFilter", []);
  const [kitFilter, setKitFilter] = useSessionState<string>("sc_kitFilter", ''); 
  
  const [flyoutOpen, setFlyoutOpen] = useState<boolean>(false);
  const [flyoutType, setFlyoutType] = useState<"product" | "material">("product");
  const [flyoutId, setFlyoutId] = useState<string>("");
  
  const [toasts, setToasts] = useState<Toast[]>([]);
  const [vPos, setVPos] = useState<VirtualPO[]>([]);
  const [initialVpoIds, setInitialVpoIds] = useState<Set<string>>(new Set());
  const [includeOtherMoNeed, setIncludeOtherMoNeed] = useState(false);
  const [planDropdownOpen, setPlanDropdownOpen] = useState(false);
  const [newPlanName, setNewPlanName] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const [isDirty, setIsDirty] = useState(false);
  const [soFilterOpen, setSoFilterOpen] = useState(false);
  const [soTypeApplied, setSoTypeApplied] = useState<string>('');
  const [soSearchQuery, setSoSearchQuery] = useState('');
  const [categoryFilterOpen, setCategoryFilterOpen] = useState(false);
  const [materialFilterOpen, setMaterialFilterOpen] = useState(false);
  const [materialSearchQuery, setMaterialSearchQuery] = useState('');
  const [kitFilterOpen, setKitFilterOpen] = useState(false);
  const [popupInfo, setPopupInfo] = useState<{ x: number; top?: number; bottom?: number; maxHeight?: number; content: React.ReactNode } | null>(null);
  const [editingVpo, setEditingVpo] = useState<VirtualPO | null>(null);
  const [vpoForm, setVpoForm] = useState({ materialId: "", qty: 0, date: "", vendor: "" });
  const [suggestBuyModal, setSuggestBuyModal] = useState<{id?: string, materialId: string, qty: number} | null>(null);

  const selectedPlan = plans.find((p) => p.id === selectedPlanId);

  const loadVposForPlan = useCallback((planId: string) => {
    if (!planVpoTable || !planVpoPlanField?.id || !planVpoMaterialField?.id || !planVpoQtyField?.id) return [];
    const planVpos: VirtualPO[] = [];
    planVpoRecords.forEach(r => {
      const pLinks = getLinks(r, planVpoPlanField);
      if (pLinks.some(l => l.id === planId)) {
        let qty = 0; const rawQty = getVal(r, planVpoQtyField);
        if (typeof rawQty === 'number') qty = rawQty;
        let vendor = 'Virtual Supplier';
        const vVal = getVal(r, planVpoVendorField);
        if (typeof vVal === 'string') vendor = vVal;
        else if (Array.isArray(vVal) && vVal.length > 0) vendor = vVal[0]?.name ?? vendor;
        
        const dateString = planVpoDateField?.id ? parseStrictDate(getVal(r, planVpoDateField)) : '';
        planVpos.push({ id: r.id, materialId: getFirstLinkId(r, planVpoMaterialField), qty, date: dateString, vendor });
      }
    });
    return planVpos;
  }, [planVpoTable, planVpoPlanField, planVpoMaterialField, planVpoQtyField, planVpoDateField, planVpoVendorField, planVpoRecords]);

  useEffect(() => {
    let cancelled = false;
    if (selectedPlanId) {
      const loaded = loadVposForPlan(selectedPlanId);
      if (!cancelled) {
        setVPos(loaded);
        setInitialVpoIds(new Set(loaded.map(v => v.id)));
      }
    } else {
      setVPos([]);
      setInitialVpoIds(new Set());
    }
    return () => { cancelled = true; };
  }, [selectedPlanId, loadVposForPlan]);

  const addToast = useCallback((type: "success" | "error", message: string) => {
    const id = uid();
    setToasts(p => [...p, { id, type, message }]);
    setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), 4000);
  }, []);

  const expandSO = useCallback((e: React.MouseEvent, soId: string) => {
    e.preventDefault(); e.stopPropagation();
    const record = soHeaderRecords.find(r => r.id === soId);
    if (record) expandRecord(record);
  }, [soHeaderRecords]);

  const expandPO = useCallback((e: React.MouseEvent, poHeaderId: string) => {
    e.preventDefault(); e.stopPropagation();
    const record = poHeaderRecords.find(r => r.id === poHeaderId);
    if (record) expandRecord(record);
  }, [poHeaderRecords]);

  const visibleOrderIds = useMemo(() => {
    if (selectedSOIds.length > 0) return selectedSOIds;
    if (selectedPlan?.linkedOrderIds && selectedPlan.linkedOrderIds.length > 0) return selectedPlan.linkedOrderIds;
    return []; 
  }, [selectedPlan, selectedSOIds]);

  const allCategories = useMemo(() => {
    const cats = new Set<string>();
    products.forEach((p) => {
      bom.filter(b => b.productId === p.id).forEach((b) => {
        const mat = materials.find(m => m.id === b.materialId);
        if (mat?.category) cats.add(mat.category);
      });
    });
    return Array.from(cats).sort();
  }, [products, bom, materials]);

  const visibleSalesOrders = useMemo(() => {
    let f = salesOrderHeaders.filter(so => visibleOrderIds.includes(so.id));
    if (dateCutoff) f = f.filter(so => so.date && so.date <= dateCutoff);
    if (categoryFilter.length > 0 && f.length > 0) {
      f = f.filter(so => {
        for (const lid of so.lineItemIds) {
          const li = lineItemDataById.get(lid);
          if (li) {
            const bes = bom.filter(b => b.productId === li.productId);
            for (const b of bes) {
              const mat = materials.find(m => m.id === b.materialId);
              if (mat && categoryFilter.includes(mat.category)) return true;
            }
          }
        }
        return false;
      });
    }
    return f;
  }, [salesOrderHeaders, visibleOrderIds, dateCutoff, categoryFilter, bom, materials, lineItemDataById]);

  const getLineItemsForOrder = useCallback((orderId: string) => {
    const order = salesOrderHeaders.find(so => so.id === orderId);
    if (!order) return [];

    const productQtyToHide = new Map<string, number>();

    for (const lid of order.lineItemIds) {
      const li = lineItemDataById.get(lid);
      if (li && li.kitId && li.isFullyDone && (li.kitQty ?? 0) > 0) {
        const kitRecord = kits.find(k => k.id === li.kitId);
        if (kitRecord) {
          kitRecord.items.forEach(ki => {
            const currentToHide = productQtyToHide.get(ki.productId) || 0;
            const qtyEatenByKit = ki.qty * (li.kitQty || 0);
            productQtyToHide.set(ki.productId, currentToHide + qtyEatenByKit);
          });
        }
      }
    }

    const items: LineItemData[] = [];
    for (const lid of order.lineItemIds) {
      const li = lineItemDataById.get(lid);
      if (!li) continue;

      if (li.isFullyDone) continue;

      let finalQty = li.qty || 0;

      if (li.productId) {
        const toHide = productQtyToHide.get(li.productId) || 0;
        if (toHide > 0) {
          if (finalQty <= toHide) {
            productQtyToHide.set(li.productId, toHide - finalQty);
            continue;
          } else {
            productQtyToHide.set(li.productId, 0);
            finalQty = finalQty - toHide;
          }
        }
      }

      const isPendingKitHeader = !li.productId && li.kitId && (li.kitQty || 0) > 0;
      
      if (finalQty > 0 || isPendingKitHeader) {
        items.push({ ...li, qty: finalQty });
      }
    }
    return items;
  }, [salesOrderHeaders, lineItemDataById, kits]);

  // materialId -> (date -> unpicked qty), open MOs only.
  // Date = earliest linked SO date (falls back to MO due date if none found).
  const stillToPickByMaterialDate = useMemo(() => {
    const map = new Map<string, Map<string, number>>();
    if (!moPickStillToPickField || !moPickMaterialField || !moPickMoField) return map;
    const moById = new Map(moRecords.map(r => [r.id, r] as [string, typeof r]));
    const soDateById = new Map(salesOrderHeaders.map(so => [so.id, so.date] as [string, string]));
    moPickRecords.forEach(line => {
      const stp = line.getCellValue(moPickStillToPickField.id);
      if (typeof stp !== 'number' || stp <= 0) return;
      const mLinks = (line.getCellValue(moPickMaterialField.id) as any[]) || [];
      if (!mLinks.length) return;
      const moLinks = (line.getCellValue(moPickMoField.id) as any[]) || [];
      const mo = moLinks.length ? moById.get(moLinks[0].id) : null;
      if (!mo) return;
      if (moStatusFieldProp) {
        const st = (mo.getCellValueAsString(moStatusFieldProp.id) || '').toLowerCase().trim();
        if (st !== 'created' && st !== 'released') return;
      }
      // Follow the SO date: earliest date among the MO's linked Sales Orders
      let date = '';
      if (moSoHeadersField) {
        const soLinks = (mo.getCellValue(moSoHeadersField.id) as any[]) || [];
        const soDates = soLinks.map(l => soDateById.get(l.id)).filter((d): d is string => !!d).sort();
        if (soDates.length) date = soDates[0];
      }
      if (!date && moDueDateField) date = parseStrictDate(mo.getCellValue(moDueDateField.id));
      if (!date) return;
      const matId = mLinks[0].id;
      if (!map.has(matId)) map.set(matId, new Map());
      const dm = map.get(matId)!;
      dm.set(date, (dm.get(date) ?? 0) + stp);
    });
    return map;
  }, [moPickRecords, moRecords, moPickStillToPickField, moPickMaterialField, moPickMoField, moStatusFieldProp, moDueDateField, moSoHeadersField, salesOrderHeaders]);
  const selectedOrderIdSet = useMemo(() => new Set(visibleSalesOrders.map(so => so.id)), [visibleSalesOrders]);

  // MO unpicked belonging to SELECTED SOs only
  const selectedOnlyStillToPickByMaterialDate = useMemo(() => {
    const map = new Map<string, Map<string, number>>();
    if (!moPickStillToPickField || !moPickMaterialField || !moPickMoField || !moSoHeadersField) return map;
    const moById = new Map(moRecords.map(r => [r.id, r] as [string, typeof r]));
    const soDateById = new Map(salesOrderHeaders.map(so => [so.id, so.date] as [string, string]));
    moPickRecords.forEach(line => {
      const stp = line.getCellValue(moPickStillToPickField.id);
      if (typeof stp !== 'number' || stp <= 0) return;
      const mLinks = (line.getCellValue(moPickMaterialField.id) as any[]) || [];
      if (!mLinks.length) return;
      const moLinks = (line.getCellValue(moPickMoField.id) as any[]) || [];
      const mo = moLinks.length ? moById.get(moLinks[0].id) : null;
      if (!mo) return;
      if (moStatusFieldProp) {
        const st = (mo.getCellValueAsString(moStatusFieldProp.id) || '').toLowerCase().trim();
        if (st !== 'created' && st !== 'released') return;
      }
      const soLinks = (mo.getCellValue(moSoHeadersField.id) as any[]) || [];
      if (!soLinks.some(l => selectedOrderIdSet.has(l.id))) return; // foreign MO -> excluded here
      let date = '';
      const soDates = soLinks.map(l => soDateById.get(l.id)).filter((d): d is string => !!d).sort();
      if (soDates.length) date = soDates[0];
      if (!date && moDueDateField) date = parseStrictDate(mo.getCellValue(moDueDateField.id));
      if (!date) return;
      const matId = mLinks[0].id;
      if (!map.has(matId)) map.set(matId, new Map());
      const dm = map.get(matId)!;
      dm.set(date, (dm.get(date) ?? 0) + stp);
    });
    return map;
  }, [moPickRecords, moRecords, moPickStillToPickField, moPickMaterialField, moPickMoField, moStatusFieldProp, moDueDateField, moSoHeadersField, salesOrderHeaders, selectedOrderIdSet]);

  // What pills, Bal, and the sandbox waterfall actually use (All always uses the full map)
  const effectiveMoNeedByMaterialDate = includeOtherMoNeed ? stillToPickByMaterialDate : selectedOnlyStillToPickByMaterialDate;
  // Per-date MO detail for the demand popup: which MO still needs how much, for which SO(s)
  const stillToPickDetailByMaterialDate = useMemo(() => {
    const map = new Map<string, Map<string, { moRecId: string; moName: string; soLabel: string; qty: number; isSelected: boolean }[]>>();
    if (!moPickStillToPickField || !moPickMaterialField || !moPickMoField) return map;
    const moById = new Map(moRecords.map(r => [r.id, r] as [string, typeof r]));
    const soDateById = new Map(salesOrderHeaders.map(so => [so.id, so.date] as [string, string]));
    moPickRecords.forEach(line => {
      const stp = line.getCellValue(moPickStillToPickField.id);
      if (typeof stp !== 'number' || stp <= 0) return;
      const mLinks = (line.getCellValue(moPickMaterialField.id) as any[]) || [];
      if (!mLinks.length) return;
      const moLinks = (line.getCellValue(moPickMoField.id) as any[]) || [];
      const mo = moLinks.length ? moById.get(moLinks[0].id) : null;
      if (!mo) return;
      if (moStatusFieldProp) {
        const st = (mo.getCellValueAsString(moStatusFieldProp.id) || '').toLowerCase().trim();
        if (st !== 'created' && st !== 'released') return;
      }
      let date = '';
      let soLabel = '';
      if (moSoHeadersField) {
        const soLinks = (mo.getCellValue(moSoHeadersField.id) as any[]) || [];
        var isSelCard = soLinks.some(l => selectedOrderIdSet.has(l.id));
        const soDates = soLinks.map(l => soDateById.get(l.id)).filter((d): d is string => !!d).sort();
        if (soDates.length) date = soDates[0];
      }
      if (!date && moDueDateField) date = parseStrictDate(mo.getCellValue(moDueDateField.id));
      if (!date) return;
      const matId = mLinks[0].id;
      if (!map.has(matId)) map.set(matId, new Map());
      const dm = map.get(matId)!;
      if (!dm.has(date)) dm.set(date, []);
      const arr = dm.get(date)!;
      const existing = arr.find(x => x.moRecId === mo.id);
      if (existing) existing.qty += stp;
      else arr.push({ moRecId: mo.id, moName: mo.name || 'MO', soLabel, qty: stp, isSelected: typeof isSelCard !== 'undefined' ? isSelCard : false });
    });
    return map;
  }, [moPickRecords, moRecords, moPickStillToPickField, moPickMaterialField, moPickMoField, moStatusFieldProp, moDueDateField, moSoHeadersField, salesOrderHeaders, selectedOrderIdSet]);
  const allDates = useMemo(() => {
    const d = new Set<string>();
    salesOrderHeaders.forEach(so => so.date && d.add(so.date));
    purchaseOrders.forEach(po => po.date && d.add(po.date));
    vPos.forEach(vpo => vpo.date && d.add(vpo.date));
    stillToPickByMaterialDate.forEach(dm => dm.forEach((_q, date) => { if (date) d.add(date); }));
    return Array.from(d).filter(x => x).sort();
  }, [salesOrderHeaders, purchaseOrders, vPos, stillToPickByMaterialDate]);

  const committedSOsSandbox = useMemo(() => visibleSalesOrders.filter(so => so.status === "Committed"), [visibleSalesOrders]);
  const timePhasedSOsSandbox = useMemo(() => visibleSalesOrders.filter(so => so.status === "TimePhasedDemand"), [visibleSalesOrders]);
  const committedSOsRealWorld = useMemo(() => salesOrderHeaders.filter(so => so.status === "Committed"), [salesOrderHeaders]);
  const timePhasedSOsRealWorld = useMemo(() => salesOrderHeaders.filter(so => so.status === "TimePhasedDemand"), [salesOrderHeaders]);

  // Uses qtyPer (with overage) — correct for planning/waterfall, unchanged
  const getMaterialDemandFromOrder = useCallback((so: SalesOrderHeader) => {
    const map = new Map<string, number>();
    const lis = getLineItemsForOrder(so.id);
    for (const li of lis) {
      bom.filter(b => b.productId === li.productId).forEach(b => {
        const d = b.qtyPer * li.qty;
        if (d > 0) map.set(b.materialId, (map.get(b.materialId) ?? 0) + d);
      });
    }
    return map;
  }, [bom, getLineItemsForOrder]);

  const buildUnreleasedDemandDated = useCallback((orders: SalesOrderHeader[]) => {
    const map = new Map<string, Map<string, number>>();
    const soliRecById = new Map(soLineItemRecords.map(r => [r.id, r] as [string, typeof r]));
    orders.forEach(so => {
      const covered = new Map<string, number>();
      for (const lid of so.lineItemIds) {
        const kitData = soLineItemKitData.get(lid);
        if (!kitData || (kitData.kitQty ?? 0) <= 0) continue;
        const rec = soliRecById.get(lid);
        if (!rec) continue;
        const status = soLineItemStatusField ? (rec.getCellValueAsString(soLineItemStatusField.id) || '').toLowerCase().trim() : '';
        const rawAlloc = soLineItemMoAllocField ? rec.getCellValue(soLineItemMoAllocField.id) : null;
        const hasAlloc = Array.isArray(rawAlloc) && rawAlloc.length > 0;
        if (status === 'released mo' || hasAlloc) {
          const kitRec = kits.find(k => k.id === kitData.kitId);
          kitRec?.items.forEach(ki => {
            covered.set(ki.productId, (covered.get(ki.productId) ?? 0) + ki.qty * kitData.kitQty);
          });
        }
      }
      for (const lid of so.lineItemIds) {
        const li = lineItemDataById.get(lid);
        if (!li || !li.productId || li.hasAllocation || li.isFullyDone) continue;
        let qty = li.qty || 0;
        const c = covered.get(li.productId) ?? 0;
        if (c > 0) {
          const eaten = Math.min(qty, c);
          covered.set(li.productId, c - eaten);
          qty -= eaten;
        }
        if (qty <= 0) continue;
        bom.filter(b => b.productId === li.productId).forEach(b => {
          if (!map.has(b.materialId)) map.set(b.materialId, new Map());
          const dm = map.get(b.materialId)!;
          dm.set(so.date, (dm.get(so.date) ?? 0) + b.qtyPer * qty);
        });
      }
    });
    return map;
  }, [soLineItemRecords, soLineItemKitData, soLineItemStatusField, soLineItemMoAllocField, lineItemDataById, kits, bom]);

  // Time-phased (non-committed) VISIBLE SOs -> feeds the pill + sandbox waterfall
  const unreleasedSODemandByMaterialDate = useMemo(
    () => buildUnreleasedDemandDated(visibleSalesOrders.filter(so => so.status !== "Committed")),
    [buildUnreleasedDemandDated, visibleSalesOrders]);

  // Time-phased ALL open SOs -> feeds the real-world waterfall (amber check)
  const realUnreleasedSODemandByMaterialDate = useMemo(
    () => buildUnreleasedDemandDated(salesOrderHeaders.filter(so => so.status !== "Committed")),
    [buildUnreleasedDemandDated, salesOrderHeaders]);

  const totalCommittedSandboxByMaterial = useMemo(() => {
    const map = new Map<string, number>();
    buildUnreleasedDemandDated(committedSOsSandbox).forEach((dm, matId) => {
      let t = 0; dm.forEach(q => { t += q; }); map.set(matId, t);
    });
    return map;
  }, [committedSOsSandbox, buildUnreleasedDemandDated]);

  const totalCommittedRealWorldByMaterial = useMemo(() => {
    const map = new Map<string, number>();
    buildUnreleasedDemandDated(committedSOsRealWorld).forEach((dm, matId) => {
      let t = 0; dm.forEach(q => { t += q; }); map.set(matId, t);
    });
    return map;
  }, [committedSOsRealWorld, buildUnreleasedDemandDated]);

  const computeRMWaterfall = useCallback((materialId: string, useAllOrders: boolean) => {
    if (!materialId) return new Map();
    const mat = materials.find(m => m.id === materialId);
    if (!mat) return new Map();
    const committedMap = useAllOrders ? totalCommittedRealWorldByMaterial : totalCommittedSandboxByMaterial;
    const timePhasedList = useAllOrders ? timePhasedSOsRealWorld : timePhasedSOsSandbox;

    let balance = mat.onHand - (committedMap.get(materialId) ?? 0);
    const result = new Map<string, { balance: number; incoming: number; outgoing: number; timePhasedDemand: number }>();

    allDates.forEach((date) => {
      let incoming = 0;
      purchaseOrders.filter(po => po.materialId === materialId && po.date === date).forEach(po => incoming += po.qty);
      vPos.filter(vpo => vpo.materialId === materialId && vpo.date === date).forEach(vpo => incoming += vpo.qty);
      // Pick-aware outgoing: unreleased SO need on its date + MO unpicked on its SO's date
      const soNeedOut = (useAllOrders ? realUnreleasedSODemandByMaterialDate : unreleasedSODemandByMaterialDate).get(materialId)?.get(date) ?? 0;
      const moNeedOut = (useAllOrders ? stillToPickByMaterialDate : effectiveMoNeedByMaterialDate).get(materialId)?.get(date) ?? 0;
      const outgoing = soNeedOut + moNeedOut;
      const timePhasedDemand = outgoing;
      balance = balance + incoming - outgoing;
      result.set(date, { balance, incoming, outgoing, timePhasedDemand });
    });
    return result;
  }, [allDates, purchaseOrders, vPos, unreleasedSODemandByMaterialDate, realUnreleasedSODemandByMaterialDate, stillToPickByMaterialDate, effectiveMoNeedByMaterialDate, totalCommittedSandboxByMaterial, totalCommittedRealWorldByMaterial, materials]);

  const selectedFilterKit = useMemo(() => kits.find(k => k.id === kitFilter) ?? null, [kits, kitFilter]);

  const kitFilterProductIds = useMemo(() => {
    if (!selectedFilterKit) return null; 
    return new Set(selectedFilterKit.productIds);
  }, [selectedFilterKit]);

  const productIds = useMemo(() => {
    const ids = new Set<string>();
    visibleSalesOrders.forEach(so => {
      getLineItemsForOrder(so.id).forEach(li => {
        if (li.productId) {
          if (!kitFilter) {
             ids.add(li.productId);
             return;
          }
          const lineItemMatchesKit = li.kitId === kitFilter;
          const kitTableMatchesProduct = kitFilterProductIds ? kitFilterProductIds.has(li.productId) : false;
          if (lineItemMatchesKit || kitTableMatchesProduct) {
            ids.add(li.productId);
          }
        }
      });
    });
    return Array.from(ids);
  }, [visibleSalesOrders, getLineItemsForOrder, kitFilter, kitFilterProductIds]);

  const materialIds = useMemo(() => {
    const set = new Set<string>();
    visibleSalesOrders.forEach(so => {
      getLineItemsForOrder(so.id).forEach(li => {
        if (li.productId) {
          let matchesFilter = false;
          if (!kitFilter) {
             matchesFilter = true;
          } else {
             const lineItemMatchesKit = li.kitId === kitFilter;
             const kitTableMatchesProduct = kitFilterProductIds ? kitFilterProductIds.has(li.productId) : false;
             matchesFilter = lineItemMatchesKit || kitTableMatchesProduct;
          }
          if (matchesFilter) {
            bom.filter(b => b.productId === li.productId).forEach(b => set.add(b.materialId));
          }
        }
      });
    });
    
    let ids = Array.from(set);
    if (strictMaterialFilter && categoryFilter.length > 0) {
      ids = ids.filter(mId => { 
        const mat = materials.find(m => m.id === mId); 
        return mat && categoryFilter.includes(mat.category); 
      });
    }
    if (materialFilter.length > 0) ids = ids.filter(mId => materialFilter.includes(mId));
    return ids;
  }, [visibleSalesOrders, getLineItemsForOrder, bom, strictMaterialFilter, categoryFilter, materials, materialFilter, kitFilter, kitFilterProductIds]);

  const allSandboxWaterfalls = useMemo(() => {
    const map = new Map<string, ReturnType<typeof computeRMWaterfall>>();
    materialIds.filter(id => !!id).forEach(id => { try { map.set(id, computeRMWaterfall(id, false)); } catch(e) {} });
    return map;
  }, [materialIds, computeRMWaterfall]);

  const allRealWorldWaterfalls = useMemo(() => {
    const map = new Map<string, ReturnType<typeof computeRMWaterfall>>();
    materialIds.filter(id => !!id).forEach(id => { try { map.set(id, computeRMWaterfall(id, true)); } catch(e) {} });
    return map;
  }, [materialIds, computeRMWaterfall]);

  const availableNowSandboxByMaterial = useMemo(() => {
    const map = new Map<string, number>();
    materials.forEach(m => map.set(m.id, m.onHand - (totalCommittedSandboxByMaterial.get(m.id) ?? 0)));
    return map;
  }, [materials, totalCommittedSandboxByMaterial]);
  // MO maps scoped/enriched for the two-balance display:
  // selected = unpicked amounts of open MOs linked to a SELECTED SO (feeds Bal)
  // requiredAll = original Required Qty of open MOs, dated (feeds the "picked" display)
  const moMapsForDisplay = useMemo(() => {
    const selected = new Map<string, Map<string, number>>();
    const requiredAll = new Map<string, Map<string, number>>();
    const requiredSelected = new Map<string, Map<string, number>>();
    if (!moPickMaterialField || !moPickMoField) return { selected, requiredAll, requiredSelected };
    const selectedIds = new Set(visibleSalesOrders.map(so => so.id));
    const moById = new Map(moRecords.map(r => [r.id, r] as [string, typeof r]));
    const soDateById = new Map(salesOrderHeaders.map(so => [so.id, so.date] as [string, string]));
    moPickRecords.forEach(line => {
      const mLinks = (line.getCellValue(moPickMaterialField.id) as any[]) || [];
      if (!mLinks.length) return;
      const moLinks = (line.getCellValue(moPickMoField.id) as any[]) || [];
      const mo = moLinks.length ? moById.get(moLinks[0].id) : null;
      if (!mo) return;
      if (moStatusFieldProp) {
        const st = (mo.getCellValueAsString(moStatusFieldProp.id) || '').toLowerCase().trim();
        if (st !== 'created' && st !== 'released') return;
      }
      let date = '';
      let isSelected = false;
      if (moSoHeadersField) {
        const soLinks = (mo.getCellValue(moSoHeadersField.id) as any[]) || [];
        isSelected = soLinks.some(l => selectedIds.has(l.id));
        const soDates = soLinks.map(l => soDateById.get(l.id)).filter((d): d is string => !!d).sort();
        if (soDates.length) date = soDates[0];
      }
      if (!date && moDueDateField) date = parseStrictDate(mo.getCellValue(moDueDateField.id));
      if (!date) return;
      const matId = mLinks[0].id;
      const stp = moPickStillToPickField ? line.getCellValue(moPickStillToPickField.id) : null;
      if (isSelected && typeof stp === 'number' && stp > 0) {
        if (!selected.has(matId)) selected.set(matId, new Map());
        const dm = selected.get(matId)!;
        dm.set(date, (dm.get(date) ?? 0) + stp);
      }
      const req = moPickQtyField ? line.getCellValue(moPickQtyField.id) : null;
      if (typeof req === 'number' && req > 0) {
        if (!requiredAll.has(matId)) requiredAll.set(matId, new Map());
        const dm = requiredAll.get(matId)!;
        dm.set(date, (dm.get(date) ?? 0) + req);
        if (isSelected) {
          if (!requiredSelected.has(matId)) requiredSelected.set(matId, new Map());
          const dms = requiredSelected.get(matId)!;
          dms.set(date, (dms.get(date) ?? 0) + req);
        }
      }
    });
    return { selected, requiredAll, requiredSelected };
  }, [moPickRecords, moRecords, moPickMaterialField, moPickMoField, moPickStillToPickField, moPickQtyField, moStatusFieldProp, moDueDateField, moSoHeadersField, salesOrderHeaders, visibleSalesOrders]);

  // Helper function to calculate unreleased SO demand for a given list of sales orders


  // Bal line: cumulative balance over SELECTED SOs (all statuses) + their MOs,
  // crediting incoming POs/VPOs dated on or before each date.
  const selectedShelfBalByMaterialDate = useMemo(() => {
    const map = new Map<string, Map<string, number>>();
    const selNeed = buildUnreleasedDemandDated(visibleSalesOrders);
    materials.forEach(m => {
      let bal = m.onHand;
      const dm = new Map<string, number>();
      allDates.forEach(date => {
        let incoming = 0;
        purchaseOrders.filter(po => po.materialId === m.id && po.date === date).forEach(po => { incoming += po.qty; });
        vPos.filter(vpo => vpo.materialId === m.id && vpo.date === date).forEach(vpo => { incoming += vpo.qty; });
        const so = selNeed.get(m.id)?.get(date) ?? 0;
        const mo = effectiveMoNeedByMaterialDate.get(m.id)?.get(date) ?? 0;
        bal = bal + incoming - so - mo;
        dm.set(date, bal);
      });
      map.set(m.id, dm);
    });
    return map;
  }, [materials, allDates, buildUnreleasedDemandDated, visibleSalesOrders, effectiveMoNeedByMaterialDate, purchaseOrders, vPos]);
  // materialId -> (SO date -> qty). ONLY demand with no MO yet.
  // Kit lines allocated to an MO (or status "Released MO") hand their products'
  // material demand to the picklist — that remainder lives in Still To Pick.


  const kitDemandMap = useMemo(() => {
    const map = new Map<string, { totalQty: number; inMoQty: number; moNames: Set<string>; byDate: Map<string, number>; inMoByDate: Map<string, number>; moNamesByDate: Map<string, Set<string>> }>();
    visibleSalesOrders.forEach(so => {
      for (const lid of so.lineItemIds) {
        const li = lineItemDataById.get(lid);
        if (!li?.kitId || !li.kitQty || li.isFullyDone) continue;
        if (!map.has(li.kitId)) map.set(li.kitId, { totalQty: 0, inMoQty: 0, moNames: new Set(), byDate: new Map(), inMoByDate: new Map(), moNamesByDate: new Map() });
        const entry = map.get(li.kitId)!;
        entry.totalQty += li.kitQty;
        entry.byDate.set(so.date, (entry.byDate.get(so.date) ?? 0) + li.kitQty);
        if (li.isInMO) {
          entry.inMoQty += li.kitQty;
          entry.inMoByDate.set(so.date, (entry.inMoByDate.get(so.date) ?? 0) + li.kitQty);
          if (!entry.moNamesByDate.has(so.date)) entry.moNamesByDate.set(so.date, new Set());
          const dateSet = entry.moNamesByDate.get(so.date)!;
          (li.moNames ?? []).forEach(n => { entry.moNames.add(n); dateSet.add(n); });
        }
      }
    });
    return map;
  }, [visibleSalesOrders, lineItemDataById]);

  const visibleKits = useMemo(() => {
    return kits.filter(k => {
      if ((kitDemandMap.get(k.id)?.totalQty ?? 0) === 0) return false;
      if (kitFilter && k.id !== kitFilter) return false; 
      return true;
    });
  }, [kits, kitDemandMap, kitFilter]);

  const computeFGBuckets = useCallback((productId: string, date: string) => {
    const bucket = { committedReady: 0, committedShort: 0, atpReady: 0, late: 0, short: 0, committedReadySOs: [] as any[], committedShortSOs: [] as any[], atpReadySOs: [] as any[], lateSOs: [] as any[], shortSOs: [] as any[] };
    const pSOs = visibleSalesOrders.filter(so => so.date === date && getLineItemsForOrder(so.id).some(li => li.productId === productId));

    pSOs.forEach((so) => {
      const lis = getLineItemsForOrder(so.id).filter(li => li.productId === productId);
      const tQty = lis.reduce((s, li) => s + li.qty, 0);
      const mNeed = new Set<string>();
      bom.filter(b => b.productId === productId).forEach(b => mNeed.add(b.materialId));

      let allPos = true; let recLat = false; let nRec = false;
      mNeed.forEach((mId) => {
        const wf = allSandboxWaterfalls.get(mId) || new Map(); 
        const tD = wf.get(date);
        if (tD && tD.balance < 0) {
          allPos = false; let rec = false;
          for (const [d, dt] of wf.entries()) { if (d > date && dt.balance >= 0) { rec = true; break; } }
          if (rec) recLat = true; else nRec = true;
        }
      });

      if (so.status === "Committed") {
        if (allPos) { bucket.committedReady += tQty; bucket.committedReadySOs.push(so); }
        else { bucket.committedShort += tQty; bucket.committedShortSOs.push(so); }
      } else {
        if (allPos) { bucket.atpReady += tQty; bucket.atpReadySOs.push(so); }
        else if (nRec) { bucket.short += tQty; bucket.shortSOs.push(so); }
        else if (recLat) { bucket.late += tQty; bucket.lateSOs.push(so); }
        else { bucket.atpReady += tQty; bucket.atpReadySOs.push(so); }
      }
    });
    return bucket;
  }, [visibleSalesOrders, allSandboxWaterfalls, bom, getLineItemsForOrder]);

  const activeDates = useMemo(() => {
    return allDates.filter(date => {
      const fga = productIds.some(pid => { const b = computeFGBuckets(pid, date); return b.committedReady>0 || b.committedShort>0 || b.atpReady>0 || b.late>0 || b.short>0; });
      const rma = materialIds.some(mId => { const w = allSandboxWaterfalls.get(mId); const need = (unreleasedSODemandByMaterialDate.get(mId)?.get(date) ?? 0) + (effectiveMoNeedByMaterialDate.get(mId)?.get(date) ?? 0); return need > 0 || (w && w.get(date) && (w.get(date)!.incoming>0 || w.get(date)!.outgoing>0)); });
      const kda = visibleKits.some(k => (kitDemandMap.get(k.id)?.byDate.get(date) ?? 0) > 0);
      return fga || rma || kda;
    });
  }, [allDates, productIds, materialIds, computeFGBuckets, allSandboxWaterfalls, visibleKits, kitDemandMap, unreleasedSODemandByMaterialDate, effectiveMoNeedByMaterialDate]);

  const handlePillClick = useCallback((e: React.MouseEvent, content: React.ReactNode) => {
    e.stopPropagation();
    const rect = (e.target as HTMLElement).getBoundingClientRect();
    const popupWidth = 420;
    const padding = 16;
    const minPopupHeight = 200;
    let x = rect.left + rect.width / 2 - popupWidth / 2;
    if (x < padding) x = padding;
    if (x + popupWidth > window.innerWidth - padding) x = window.innerWidth - popupWidth - padding;
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    let top: number | undefined;
    let bottom: number | undefined;
    let maxHeight: number;
    if (spaceBelow > minPopupHeight + 40) {
      top = rect.bottom + 8;
      maxHeight = Math.min(520, spaceBelow - 30);
    } else if (spaceAbove > minPopupHeight + 40) {
      bottom = window.innerHeight - rect.top + 8;
      maxHeight = Math.min(520, spaceAbove - 30);
    } else {
      top = 40;
      maxHeight = window.innerHeight - 80;
    }
    setPopupInfo({ x, top, bottom, maxHeight, content });
  }, []);

  const closePopup = useCallback(() => setPopupInfo(null), []);
  const openFlyout = useCallback((type: "product" | "material", id: string) => { setFlyoutType(type); setFlyoutId(id); setFlyoutOpen(true); }, [setFlyoutType, setFlyoutId, setFlyoutOpen]);
  const closeFlyout = useCallback(() => { setFlyoutOpen(false); }, [setFlyoutOpen]);

  const generateMakeOrder = async (
  kit: Kit, 
  totalQty: number, 
  relatedSOs: SalesOrderHeader[], 
  targetDate?: string
) => {
  if (!moTableProp || !moKitField || !moAllocTableProp || !moAllocMoField || 
      !moAllocSoLineField || !moAllocQtyField || !moPickTableProp || 
      !moPickTypeField || !moPickQtyField) {
    addToast("error", "Make Order tables/fields are not fully mapped in settings.");
    return;
  }

  // --- Early permission checks ---
  if (!moTableProp.hasPermissionToCreateRecords()) {
    addToast("error", `Permission Denied: Cannot create records in '${moTableProp.name}' table.`);
    return;
  }
  if (!moAllocTableProp.hasPermissionToCreateRecords()) {
    addToast("error", `Permission Denied: Cannot create records in '${moAllocTableProp.name}' table.`);
    return;
  }
  if (!moPickTableProp.hasPermissionToCreateRecords()) {
    addToast("error", `Permission Denied: Cannot create records in '${moPickTableProp.name}' table.`);
    return;
  }

  setIsSaving(true);
  let currentStep = "starting process";

  try {
    // ============================================
    // 1. Calculate Due Date (with safe fallback)
    // ============================================
    currentStep = "calculating due date";
    let dueDate = targetDate || "";
    
    if (!dueDate && relatedSOs.length > 0) {
      const validDates = relatedSOs
        .map(so => new Date(so.date).getTime())
        .filter(t => !isNaN(t));
      
      if (validDates.length > 0) {
        dueDate = new Date(Math.min(...validDates)).toISOString().split('T')[0];
      }
    }
    
    // Final fallback: today's date (prevents "required field" errors)
    if (!dueDate) {
      dueDate = new Date().toISOString().split('T')[0];
    }

    // ============================================
    // 2. Collect which SOs and SOLIs this MO fulfills
    // ============================================
    currentStep = "collecting linked orders";
    const soHeaderIdsForMo = new Set<string>();
    const soLineItemIdsForMo = new Set<string>();

    relatedSOs.forEach(so => {
      const lines = getLineItemsForOrder(so.id);
      lines.forEach(li => {
        if (li.kitId === kit.id && (li.kitQty ?? 0) > 0 && !li.hasAllocation) {
          soHeaderIdsForMo.add(so.id);
          soLineItemIdsForMo.add(li.id);
        }
      });
    });

    // ============================================
    // 3. Build MO Header fields (SMART link handling)
    // ============================================
    currentStep = "building Make Order header";
    
    const moFields: any = {
      [moKitField.id]: [{ id: kit.id }]
    };

    if (moDueDateField) {
      moFields[moDueDateField.id] = dueDate;
    }

    // Smart link helper: respects single vs multiple link fields
    const getLinkValue = (field: Field | undefined, idSet: Set<string>) => {
      if (!field || idSet.size === 0) return undefined;
      const arr = Array.from(idSet).map(id => ({ id }));
      const allowsMultiple = field.options?.multipleRecordLinks !== false;
      return allowsMultiple ? arr : arr[0]; // single link → object, multiple → array
    };

    if (moSoHeadersField) {
      const linkVal = getLinkValue(moSoHeadersField, soHeaderIdsForMo);
      if (linkVal) moFields[moSoHeadersField.id] = linkVal;
    }

    if (moSoLineItemsField) {
      const linkVal = getLinkValue(moSoLineItemsField, soLineItemIdsForMo);
      if (linkVal) moFields[moSoLineItemsField.id] = linkVal;
    }

    // Permission check with actual fields
    const headerPerm = moTableProp.checkPermissionsForCreateRecord(moFields);
    if (!headerPerm.hasPermission) {
      const reason = headerPerm.reasonDisplayString ?? "Unknown reason";
      addToast("error", `MO Header blocked: ${reason}`);
      console.warn("[MO Header] create blocked →", headerPerm);
      setIsSaving(false);
      return;
    }

    // Create the Make Order header
    currentStep = "creating Make Order header in Airtable";
    const moId = await moTableProp.createRecordAsync(moFields);
    console.log(`[MO] Created header ${moId} for kit "${kit.name}"`);

    // ============================================
    // 4. Create MO Allocations (one per SOLI)
    // ============================================
    currentStep = "creating MO Allocations";
    const allocRecords: any[] = [];

    relatedSOs.forEach(so => {
      const lines = getLineItemsForOrder(so.id);
      lines.forEach(li => {
        if (li.kitId === kit.id && (li.kitQty ?? 0) > 0 && !li.hasAllocation) {
          const allocFields: any = {
            [moAllocMoField.id]: [{ id: moId }],
            [moAllocSoLineField.id]: [{ id: li.id }],
            [moAllocQtyField.id]: li.kitQty
          };

          if (moAllocSoHeaderField) {
            allocFields[moAllocSoHeaderField.id] = [{ id: so.id }];
          }

          allocRecords.push({ fields: allocFields });
        }
      });
    });

    if (allocRecords.length > 0) {
      const allocPerm = moAllocTableProp.checkPermissionsForCreateRecords();
      if (!allocPerm.hasPermission) {
        addToast("error", `MO Allocations blocked: ${allocPerm.reasonDisplayString ?? "no permission"}`);
        console.warn("[MO Allocations] blocked →", allocPerm);
        setIsSaving(false);
        return;
      }

      for (let i = 0; i < allocRecords.length; i += 50) {
        await moAllocTableProp.createRecordsAsync(allocRecords.slice(i, i + 50));
      }
      console.log(`[MO] Created ${allocRecords.length} allocation records`);
    } else {
      console.log(`[MO] No allocations created (no matching unallocated lines)`);
    }

    // ============================================
    // 5. Explode BOM → Create Picklist records
    // ============================================
    currentStep = "creating MO Picklist";
    const picklistMap = new Map<string, { type: string; pId?: string; mId?: string; qty: number; qtyPerKit: number }>();

    kit.items.forEach(ki => {
      const reqProductQty = ki.qty * totalQty;

      // Product line
      const pKey = 'p_' + ki.productId;
      if (!picklistMap.has(pKey)) {
        picklistMap.set(pKey, { type: 'Product', pId: ki.productId, qty: 0, qtyPerKit: 0 });
      }
      const pEntry = picklistMap.get(pKey)!;
      pEntry.qty += reqProductQty;
      pEntry.qtyPerKit += ki.qty;

      // Raw materials from BOM
      bom.filter(b => b.productId === ki.productId).forEach(b => {
        const perKit = (b.moQtyPer ?? b.qtyPer) * ki.qty;
        const mKey = 'm_' + b.materialId;
        if (!picklistMap.has(mKey)) {
          picklistMap.set(mKey, { type: 'Raw Material', mId: b.materialId, qty: 0, qtyPerKit: 0 });
        }
        const mEntry = picklistMap.get(mKey)!;
        mEntry.qty += perKit * totalQty;
        mEntry.qtyPerKit += perKit;
      });
    });

    // Every picklist line links the SOLIs this MO fulfills -> feeds QTY Demand rollup (live)
    const soliLinks = Array.from(soLineItemIdsForMo).map(id => ({ id }));

    const pickRecords: any[] = [];
    picklistMap.forEach(item => {
      const typeVal = moPickTypeField.type === 'singleSelect'
        ? { name: item.type }
        : item.type;

      const pFields: any = {
        [moPickMoField.id]: [{ id: moId }],
        [moPickTypeField.id]: typeVal,
        [moPickQtyField.id]: item.qty
      };

      if (moPickSoLineField && soliLinks.length > 0) {
        pFields[moPickSoLineField.id] = soliLinks;
      }
      if (moPickQtyPerKitField) {
        pFields[moPickQtyPerKitField.id] = item.qtyPerKit;
      }

      if (item.type === 'Product' && moPickProductField && item.pId) {
        pFields[moPickProductField.id] = [{ id: item.pId }];
      }
      if (item.type === 'Raw Material' && moPickMaterialField && item.mId) {
        pFields[moPickMaterialField.id] = [{ id: item.mId }];
      }

      pickRecords.push({ fields: pFields });
    });

    if (pickRecords.length > 0) {
      const pickPerm = moPickTableProp.checkPermissionsForCreateRecords();
      if (!pickPerm.hasPermission) {
        addToast("error", `MO Picklist blocked: ${pickPerm.reasonDisplayString ?? "no permission"}`);
        console.warn("[MO Picklist] blocked →", pickPerm);
        setIsSaving(false);
        return;
      }

      for (let i = 0; i < pickRecords.length; i += 50) {
        await moPickTableProp.createRecordsAsync(pickRecords.slice(i, i + 50));
      }
      console.log(`[MO] Created ${pickRecords.length} picklist records`);
    } else {
      console.log(`[MO] No picklist records created`);
    }
    // ============================================
    // 6. Flip SOLI Status -> "Released MO"
    // Prevents double-counting in Materials "Real Qty Needed"
    // (line must leave "Unreleased SO Demand" once it has a picklist)
    // ============================================
    currentStep = "updating SOLI statuses to Released MO";
    if (soLineItemsTable && soLineItemStatusField && soLineItemIdsForMo.size > 0) {
      const statusChoices = (soLineItemStatusField.options?.choices as any[]) || [];
      const releasedChoice = statusChoices.find(
        c => c.name.toLowerCase().trim() === 'released mo'
      );
      if (!releasedChoice) {
        console.warn('[MO] "Released MO" option not found on SOLI Status — statuses NOT updated');
      } else if (!soLineItemsTable.hasPermissionToUpdateRecords()) {
        console.warn('[MO] No permission to update SO Line Items — statuses NOT updated');
      } else {
        const statusUpdates = Array.from(soLineItemIdsForMo).map(id => ({
          id,
          fields: { [soLineItemStatusField.id]: { id: releasedChoice.id } }
        }));
        for (let i = 0; i < statusUpdates.length; i += 50) {
          await soLineItemsTable.updateRecordsAsync(statusUpdates.slice(i, i + 50));
        }
        console.log(`[MO] Updated ${statusUpdates.length} SOLI statuses to Released MO`);
      }
    }

    // ============================================
    // SUCCESS
    // ============================================
    addToast("success", `Make Order created successfully for ${totalQty.toLocaleString()} units of "${kit.name}"`);
    closePopup();

  } catch (error) {
    console.error(`CRASH LOG: Failed during step -> ${currentStep}`, error);
    const errMsg = error instanceof Error ? error.message : String(error);
    addToast("error", `Failed while ${currentStep}. Airtable says: ${errMsg}`);
} finally {
    setIsSaving(false);
  }
};

  const renderKitDemandPopup = (kit: Kit, headerQty: number, targetDate?: string) => {
    const contrib: { so: SalesOrderHeader; li: LineItemData }[] = [];
    visibleSalesOrders.forEach(so => {
      if (targetDate && so.date !== targetDate) return;
      getLineItemsForOrder(so.id).forEach(li => {
        if (li.kitId === kit.id && (li.kitQty ?? 0) > 0) contrib.push({ so, li });
      });
    });

    const genLines = contrib.filter(x => !x.li.isInMO);
    const generatableQty = genLines.reduce((s, x) => s + (x.li.kitQty ?? 0), 0);
    const relatedSOs = Array.from(new Map(genLines.map(x => [x.so.id, x.so])).values());

    return (
      <div className="text-xs w-full flex flex-col">
        <div className="font-black uppercase tracking-widest px-3 py-2 rounded-lg sticky top-0 z-10 border text-purple-800 dark:text-purple-300 bg-purple-100 border-purple-200 dark:bg-purple-900/80 dark:border-purple-700 mb-2 shadow-sm backdrop-blur-md bg-opacity-95">
          {kit.name} — {targetDate ? `${formatDate(targetDate)} · ` : 'TOTAL DEMAND: '}{Math.round(headerQty).toLocaleString()} kits
        </div>
        <div className="flex-1 overflow-y-auto px-1 pb-4">
          {contrib.map(({ so, li }, idx) => {
            const inMO = li.isInMO;
            return (
              <div key={li.id + '_' + idx} className={`w-full text-left p-3 rounded-xl border ${borderColor} ${isDark ? 'bg-gray-800' : 'bg-white'} shadow-sm block mb-2 ${inMO ? 'opacity-50' : ''}`}>
                <div className="flex justify-between items-start">
                  <div className="flex flex-col">
                    <span className="font-bold text-sm uppercase text-slate-900 dark:text-slate-100">{so.customer}</span>
                    <span className={`text-[12px] ${textSecondary} uppercase tracking-wider`}>{so.name}{so.customerPO ? ` • PO: ${so.customerPO}` : ''} • {formatDate(so.date)}</span>
                    {inMO && (
                      <span className="mt-1 self-start inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[9px] font-black uppercase tracking-widest bg-indigo-100 text-indigo-700 border border-indigo-200 dark:bg-indigo-900/50 dark:text-indigo-400 dark:border-indigo-800">
                        In MO{li.moNames && li.moNames.length > 0 ? `: ${li.moNames.join(', ')}` : ''}
                      </span>
                    )}
                  </div>
                  <span className={`font-mono font-black px-2 py-0.5 rounded border ${inMO ? 'text-slate-500 bg-slate-50 border-slate-200 dark:bg-gray-900/30 dark:text-slate-500 dark:border-gray-700' : 'text-purple-600 bg-purple-50 dark:bg-purple-900/30 dark:text-purple-400 border-purple-100 dark:border-purple-800'}`}>
                    {Math.round(li.kitQty ?? 0).toLocaleString()} KITS
                  </span>
                </div>
              </div>
            );
          })}
        </div>
        <div className={`p-3 border-t ${borderColor} ${isDark?'bg-gray-800':'bg-slate-50'} sticky bottom-0 rounded-b-xl shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.1)]`}>
          <button
            onClick={() => generateMakeOrder(kit, generatableQty, relatedSOs, targetDate)}
            disabled={isSaving || generatableQty <= 0}
            className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-black uppercase tracking-widest bg-purple-600 text-white hover:bg-purple-700 transition-colors disabled:opacity-50 shadow-md hover:shadow-lg"
          >
            {isSaving ? "Processing..." : generatableQty <= 0 ? "All lines already in MO" : <><GearIcon weight="bold" size={20} /> Generate Make Order ({Math.round(generatableQty).toLocaleString()})</>}
          </button>
        </div>
      </div>
    );
  };

  const addVirtualPO = useCallback(() => {
    setEditingVpo({ id: "", materialId: materials[0]?.id ?? "", qty: 100, date: allDates[0] ?? "", vendor: "Virtual Supplier" });
    setVpoForm({ materialId: materials[0]?.id ?? "", qty: 100, date: allDates[0] ?? "", vendor: "Virtual Supplier" });
    setIsDirty(true);
  }, [allDates, materials]);

  const saveVpo = useCallback(() => {
    const newVpo: VirtualPO = { id: editingVpo?.id || uid(), materialId: vpoForm.materialId, qty: vpoForm.qty, date: vpoForm.date, vendor: vpoForm.vendor };
    setVPos(p => {
      const e = p.find(v => v.id === newVpo.id);
      return e ? p.map(v => v.id === newVpo.id ? newVpo : v) : [...p, newVpo];
    });
    setEditingVpo(null); setIsDirty(true); addToast("success", "Virtual PO saved");
  }, [editingVpo, vpoForm, addToast]);

  const deleteVpo = useCallback((id: string) => { setVPos(p => p.filter(v => v.id !== id)); setIsDirty(true); addToast("success", "Virtual PO deleted"); }, [addToast]);
  const handleSOSelectionChange = useCallback((ids: string[]) => { setSelectedSOIds(ids); setIsDirty(true); }, [setSelectedSOIds]);

  const saveVposForPlan = useCallback(async (planId: string, currentVpos: VirtualPO[]) => {
    if (!planVpoTable || !planVpoPlanField?.id || !planVpoMaterialField?.id || !planVpoQtyField?.id) return;
    const exIds = new Set<string>(initialVpoIds);
    const cIds = new Set(currentVpos.filter(v => exIds.has(v.id)).map(v => v.id));
    const toDel = Array.from(exIds).filter(id => !cIds.has(id));
    const toUp = currentVpos.filter(v => exIds.has(v.id));
    const toCr = currentVpos.filter(v => !exIds.has(v.id));

    if (toDel.length > 0 && planVpoTable.hasPermissionToDeleteRecords()) {
      for (let i=0; i<toDel.length; i+=50) await planVpoTable.deleteRecordsAsync(toDel.slice(i, i+50));
    }
    if (toUp.length > 0 && planVpoTable.hasPermissionToUpdateRecords()) {
      for (let i=0; i<toUp.length; i+=50) {
        await planVpoTable.updateRecordsAsync(toUp.slice(i, i+50).map(v => {
          const f: any = { [planVpoMaterialField.id]: [{id: v.materialId}], [planVpoQtyField.id]: v.qty };
          if (planVpoDateField?.id && v.date) f[planVpoDateField.id] = v.date;
          if (planVpoVendorField?.id) f[planVpoVendorField.id] = v.vendor;
          return { id: v.id, fields: f };
        }));
      }
    }
    if (toCr.length > 0 && planVpoTable.hasPermissionToCreateRecords()) {
      for (let i=0; i<toCr.length; i+=50) {
        await planVpoTable.createRecordsAsync(toCr.slice(i, i+50).map(v => {
          const f: any = { [planVpoPlanField.id]: [{id: planId}], [planVpoMaterialField.id]: [{id: v.materialId}], [planVpoQtyField.id]: v.qty };
          if (planVpoDateField?.id && v.date) f[planVpoDateField.id] = v.date;
          if (planVpoVendorField?.id) f[planVpoVendorField.id] = v.vendor;
          return { fields: f };
        }));
      }
    }
  }, [planVpoTable, planVpoPlanField, planVpoMaterialField, planVpoQtyField, planVpoDateField, planVpoVendorField, initialVpoIds]);

  const saveNewPlan = useCallback(async () => {
    if (!plansTable || !planNameField?.id || !newPlanName.trim()) return addToast("error", "Please enter a plan name");
    if (!plansTable.hasPermissionToCreateRecords()) return addToast("error", "No permission to create plans");
    setIsSaving(true);
    try {
      const f: any = { [planNameField.id]: newPlanName.trim() };
      if (planOrdersLinkField?.id && selectedSOIds.length > 0) f[planOrdersLinkField.id] = selectedSOIds.map(id => ({id}));
      const nId = await plansTable.createRecordAsync(f);
      if (vPos.length > 0) await saveVposForPlan(nId, vPos);
      setSelectedPlanId(nId); setNewPlanName(""); setIsDirty(false); setInitialVpoIds(new Set()); addToast("success", "Plan saved");
    } catch (e) { addToast("error", "Save failed"); } finally { setIsSaving(false); }
  }, [plansTable, planNameField, planOrdersLinkField, newPlanName, selectedSOIds, vPos, saveVposForPlan, addToast, setSelectedPlanId]);

  const updateExistingPlan = useCallback(async () => {
    if (!plansTable || !selectedPlanId) return addToast("error", "No plan selected");
    if (!plansTable.hasPermissionToUpdateRecords()) return addToast("error", "No permission");
    setIsSaving(true);
    try {
      const f: any = {};
      if (planOrdersLinkField?.id) f[planOrdersLinkField.id] = selectedSOIds.map(id => ({id}));
      await plansTable.updateRecordAsync(selectedPlanId, f);
      await saveVposForPlan(selectedPlanId, vPos);
      setIsDirty(false); setInitialVpoIds(new Set(vPos.map(v=>v.id))); addToast("success", "Plan updated");
    } catch (e) { addToast("error", "Update failed"); } finally { setIsSaving(false); }
  }, [plansTable, selectedPlanId, planOrdersLinkField, selectedSOIds, vPos, saveVposForPlan, addToast]);

  const submitBuySuggestion = useCallback(async () => {
    if (!suggestBuyModal) return;
    if (!suggestedTable || !suggestedQtyField?.id) return addToast('error', 'Suggested Buys table is not fully configured.');
    if (suggestBuyModal.id && !suggestedTable.hasPermissionToUpdateRecords()) return addToast('error', 'No permission to update.');
    if (!suggestBuyModal.id && !suggestedTable.hasPermissionToCreateRecords()) return addToast('error', 'No permission to create.');
    try {
      if (suggestBuyModal.id) {
        await suggestedTable.updateRecordAsync(suggestBuyModal.id, { [suggestedQtyField.id]: suggestBuyModal.qty });
        addToast('success', 'Suggestion updated!');
      } else {
        if (!suggestedMaterialField?.id) return addToast('error', 'Material field not configured.');
        const f: any = { [suggestedMaterialField.id]: [{ id: suggestBuyModal.materialId }], [suggestedQtyField.id]: suggestBuyModal.qty };
        if (selectedPlanId && suggestedPlanField?.id) f[suggestedPlanField.id] = [{ id: selectedPlanId }];
        await suggestedTable.createRecordAsync(f);
        addToast('success', 'Buy suggestion sent!');
      }
      setSuggestBuyModal(null);
    } catch (e) { addToast('error', 'Failed to send suggestion'); }
  }, [suggestBuyModal, suggestedTable, suggestedMaterialField, suggestedQtyField, suggestedPlanField, selectedPlanId, addToast]);

  const deletePlanById = useCallback(async (planId: string) => {
    if (!plansTable) return addToast("error", "Plans table not configured");
    if (!plansTable.hasPermissionToDeleteRecords()) return addToast("error", "No permission to delete plans");
    setIsSaving(true);
    try {
      await plansTable.deleteRecordAsync(planId);
      if (selectedPlanId === planId) { setSelectedPlanId(""); setSelectedSOIds([]); setIsDirty(false); }
      addToast("success", `Plan deleted`);
    } catch (e) { addToast("error", "Delete failed"); } finally { setIsSaving(false); }
  }, [plansTable, selectedPlanId, addToast, setSelectedPlanId, setSelectedSOIds]);

  const flyoutProduct = products.find(p => p.id === flyoutId);
  const flyoutMaterial = materials.find(m => m.id === flyoutId);

  const flyoutDemand = useMemo(() => {
    if (flyoutType === "product" && flyoutProduct) {
      let t = 0; visibleSalesOrders.forEach(so => getLineItemsForOrder(so.id).filter(li => li.productId === flyoutProduct.id).forEach(li => t += li.qty)); return t;
    }
    if (flyoutType === "material" && flyoutMaterial) {
      let t = 0;
      visibleSalesOrders.forEach(so => {
        const demandMap = getMaterialDemandFromOrder(so);
        t += (demandMap.get(flyoutMaterial.id) ?? 0);
      });
      return t;
    }
    return 0;
  }, [flyoutType, flyoutProduct, flyoutMaterial, visibleSalesOrders, getMaterialDemandFromOrder, getLineItemsForOrder]);

  const flyoutSupply = useMemo(() => {
    if (flyoutType === "material" && flyoutMaterial) return flyoutMaterial.onHand + purchaseOrders.filter(po => po.materialId === flyoutMaterial.id).reduce((s, po) => s + po.qty, 0) + vPos.filter(vpo => vpo.materialId === flyoutMaterial.id).reduce((s, v) => s + v.qty, 0);
    return 0;
  }, [flyoutType, flyoutMaterial, purchaseOrders, vPos]);

  const flyoutSOs = useMemo(() => {
    let sos: SalesOrderHeader[] = [];
    if (flyoutType === "product" && flyoutProduct) sos = visibleSalesOrders.filter(so => getLineItemsForOrder(so.id).some(li => li.productId === flyoutProduct.id));
    if (flyoutType === "material" && flyoutMaterial) {
      const pIds = bom.filter(b => b.materialId === flyoutMaterial.id).map(b => b.productId);
      sos = visibleSalesOrders.filter(so => getLineItemsForOrder(so.id).some(li => pIds.includes(li.productId)));
    }
    return sos.sort((a, b) => (new Date(a.date).getTime() || 0) - (new Date(b.date).getTime() || 0));
  }, [flyoutType, flyoutProduct, flyoutMaterial, visibleSalesOrders, bom, getLineItemsForOrder]);

  const soStatusOptions = useMemo(() => {
    if (!soOrderStatusField) return ['Pending', 'Planned', 'Committed', 'Shipped'];
    const c = soOrderStatusField.options?.choices as any[] | undefined;
    return (c && c.length > 0) ? c.map(x => x.name) : ['Pending', 'Planned', 'Committed', 'Shipped'];
  }, [soOrderStatusField]);

  const updateSOStatus = useCallback(async (soId: string, newStatus: string) => {
    if (!soHeadersTable || !soOrderStatusField) return addToast('error', 'Status field not configured');
    if (!soHeadersTable.hasPermissionToUpdateRecords()) return addToast('error', 'No permission');
    try {
      const c = soOrderStatusField.options?.choices as any[] | undefined;
      const m = c?.find(x => x.name.toLowerCase().trim() === newStatus.toLowerCase().trim());
      if (!m) return addToast('error', 'Status not found');
      await soHeadersTable.updateRecordAsync(soId, { [soOrderStatusField.id]: { id: m.id } });
      addToast('success', `Status updated`);
    } catch (e) { addToast('error', 'Update failed'); }
  }, [soHeadersTable, soOrderStatusField, addToast]);

  const flyoutPOs = useMemo(() => flyoutType === "material" && flyoutMaterial ? purchaseOrders.filter(po => po.materialId === flyoutMaterial.id).sort((a, b) => (new Date(a.date).getTime() || 0) - (new Date(b.date).getTime() || 0)) : [], [flyoutType, flyoutMaterial, purchaseOrders]);
  
  const poGroups = useMemo(() => {
    const groups = new Map();
    flyoutPOs.forEach(po => {
        if (!groups.has(po.poLineId)) groups.set(po.poLineId, { poLineId: po.poLineId, headerId: po.headerId, name: po.name, vendor: po.vendor, status: po.status, remaining: null, shipments: [] });
        if (po.isShipment) groups.get(po.poLineId).shipments.push(po);
        else groups.get(po.poLineId).remaining = po;
    });
    return Array.from(groups.values());
  }, [flyoutPOs]);

  const flyoutVPOs = useMemo(() => flyoutType === "material" && flyoutMaterial ? vPos.filter(vpo => vpo.materialId === flyoutMaterial.id).sort((a, b) => (new Date(a.date).getTime() || 0) - (new Date(b.date).getTime() || 0)) : [], [flyoutType, flyoutMaterial, vPos]);

  // STYLES
  const bgMain = isDark ? "bg-gray-800" : "bg-gray-50";
  const bgCard = isDark ? "bg-gray-700" : "bg-white";
  const borderColor = isDark ? "border-gray-600" : "border-gray-200";
  const textPrimary = isDark ? "text-gray-100" : "text-gray-900";
  const textSecondary = isDark ? "text-gray-400" : "text-gray-500";
  const hoverBgButton = isDark ? "hover:bg-gray-600" : "hover:bg-gray-100";
  const theadBg = isDark ? "bg-gray-600" : "bg-gray-100";

  const clsStrictBg = isDark ? 'border-gray-600 bg-indigo-900/30' : 'border-gray-200 bg-indigo-50/50';
  const clsKitsBadge = isDark ? 'bg-purple-900/50 text-purple-400' : 'bg-purple-100 text-purple-700';
  const clsFlyoutSubBg = isDark ? 'bg-gray-800/30' : 'bg-slate-50';
  const clsFlyoutRowBg = isDark ? 'bg-gray-800/80' : 'bg-slate-50';
  const clsSuggestCard = isDark ? 'border-indigo-800 bg-indigo-900/30' : 'border-indigo-100 bg-indigo-50/80';
  const clsPoStatusBadge = isDark ? 'bg-blue-900/50 text-blue-400 border-blue-800' : 'bg-blue-50 text-blue-700 border-blue-200';
  const clsShipmentsBg = isDark ? 'bg-gray-900/40' : 'bg-white';
  const clsVpoBorder = isDark ? 'border-purple-800 bg-purple-900/20' : 'border-purple-200 bg-purple-50';
  const clsVpoEditBtn = isDark ? 'text-purple-400 hover:bg-purple-900/50' : 'text-purple-500 hover:bg-purple-200';
  const clsVpoDeleteBtn = isDark ? 'text-rose-400 hover:bg-rose-900/50' : 'text-rose-500 hover:bg-rose-200';
  const clsVpoDivider = isDark ? 'border-purple-800/50' : 'border-purple-200';
  const clsVpoQtyBadge = isDark ? 'text-purple-300 bg-purple-900/50' : 'text-purple-800 bg-purple-200';

  if (missingConfigs.length > 0) return <ConfigurationError missingConfigs={missingConfigs} />;

  return (
    <div className={`min-h-screen w-full flex flex-col ${isDark ? 'bg-gray-800 text-gray-100' : 'bg-gray-50 text-gray-900'}`}>
      <div className="fixed top-4 right-4 z-[999] flex flex-col gap-2 pointer-events-none">
        {toasts.map((toast) => (
          <div key={toast.id} className={`px-4 py-3 rounded-md shadow-lg text-sm font-medium flex items-center gap-2 transition-all pointer-events-auto ${toast.type === "success" ? "bg-emerald-600 text-white" : "bg-rose-600 text-white"}`}>
            {toast.type === "success" ? <CheckIcon weight="bold" /> : <WarningIcon weight="bold" />} {toast.message}
          </div>
        ))}
      </div>

      <header className={`sticky top-0 z-30 ${isDark?'bg-gray-700 border-gray-600':'bg-white border-gray-200'} border-b px-4 py-3 flex flex-col gap-3 shadow-sm`}>
        <div className="flex items-center gap-2">
          <PackageIcon weight="duotone" className="w-6 h-6 text-indigo-500" />
          <span className="font-semibold text-base">Supply Chain Flow Planner</span>
          <div className="ml-auto">
            <button onClick={addVirtualPO} className="flex items-center gap-1 px-3 py-1.5 text-xs font-bold rounded-md bg-indigo-600 text-white hover:bg-indigo-700 shadow-xs">
              <PlusIcon weight="bold" /> Simulate Buy
            </button>
          </div>
        </div>

        <div className="flex items-center gap-3 w-full">
          <div className="flex items-center gap-3 flex-wrap">
          
          {/* 1. PLAN SELECTOR */}
          <div className="flex items-center gap-2 relative">
            <div className="relative">
              <button onClick={() => setPlanDropdownOpen(!planDropdownOpen)} className={`flex items-center gap-2 px-3 py-1.5 text-xs font-medium rounded-md shadow-xs border ${isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600':'border-gray-200 bg-white hover:bg-gray-100'} min-w-[140px] md:min-w-[180px] focus:ring-2 focus:ring-indigo-500`}>
                <PackageIcon weight="bold" className="text-indigo-500" />
                <span className="truncate flex-1 text-left">{selectedPlan ? selectedPlan.name : "Unsaved View"}</span>
                {selectedPlan?.status === 'Finalized' && <CheckIcon weight="bold" className="text-emerald-500" />}
                {isDirty && selectedPlanId && <CircleIcon weight="fill" className="w-2 h-2 text-amber-500" />}
                <CaretRightIcon className={`w-3.5 h-3.5 transition-transform ${planDropdownOpen ? "rotate-90" : ""}`} />
              </button>
              {planDropdownOpen && (
                <>
                  <div className="fixed inset-0 z-30" onClick={() => setPlanDropdownOpen(false)} />
                  <div className={`absolute left-0 top-full mt-1 w-72 max-h-96 overflow-y-auto rounded-lg shadow-lg border ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} z-40`}>
                    <div className={`sticky top-0 px-3 py-2 border-b ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'}`}>
                      <div className="flex items-center gap-2">
                        <input type="text" placeholder="New plan name..." value={newPlanName} onChange={(e) => setNewPlanName(e.target.value)} className={`flex-1 px-2 py-1.5 text-xs rounded-md border focus:ring-indigo-500 ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'}`} onClick={(e) => e.stopPropagation()} />
                        <button onClick={(e) => { e.stopPropagation(); if (newPlanName.trim()) { saveNewPlan(); setPlanDropdownOpen(false); } }} disabled={!newPlanName.trim() || isSaving} className="p-1.5 rounded-md bg-indigo-600 text-white hover:bg-indigo-700 disabled:opacity-50"><PlusIcon weight="bold" /></button>
                      </div>
                      {selectedSOIds.length > 0 && <div className={`text-xs ${isDark?'text-gray-400':'text-gray-500'} mt-1`}>Will include {selectedSOIds.length.toLocaleString()} orders</div>}
                    </div>
                    <div className="p-1">
                      <button onClick={() => { setSelectedPlanId(""); setSelectedSOIds([]); setIsDirty(false); setPlanDropdownOpen(false); }} className={`w-full flex items-center gap-2 px-3 py-2 rounded-md text-left ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${!selectedPlanId ? (isDark?'bg-gray-600':'bg-gray-100') : ""}`}>
                        <PackageIcon weight="regular" className="text-slate-400" /> <span className="text-sm font-medium">Clear / Unsaved View</span>
                      </button>
                      {plans.map((p) => (
                        <div key={p.id} className={`flex items-center gap-2 px-3 py-2 rounded-md ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${selectedPlanId === p.id ? (isDark?'bg-gray-600':'bg-gray-100') : ""}`}>
                          <button onClick={() => { setSelectedPlanId(p.id); setSelectedSOIds(p.linkedOrderIds); setIsDirty(false); setPlanDropdownOpen(false); }} className="flex-1 flex items-center gap-2 text-left">
                            <PackageIcon weight={selectedPlanId === p.id ? "fill" : "regular"} className="text-indigo-500" />
                            <div className="flex-1 min-w-0">
                              <div className="text-sm font-medium truncate">{p.name}</div>
                              <div className={`text-xs ${isDark?'text-gray-400':'text-gray-500'}`}>{p.linkedOrderIds.length.toLocaleString()} orders {p.status === 'Finalized' && <span className="text-emerald-500">✓ Finalized</span>}</div>
                            </div>
                          </button>
                          <button onClick={(e) => { e.stopPropagation(); deletePlanById(p.id); setPlanDropdownOpen(false); }} disabled={isSaving} className="p-1 rounded hover:bg-rose-100 text-rose-500"><TrashIcon weight="bold" /></button>
                        </div>
                      ))}
                    </div>
                  </div>
                </>
              )}
            </div>
            {isDirty && selectedPlanId && (
              <button onClick={updateExistingPlan} disabled={isSaving} className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-md shadow-xs bg-emerald-600 text-white hover:bg-emerald-700 disabled:opacity-50">
                <CheckIcon weight="bold" /> {isSaving ? "..." : "Save"}
              </button>
            )}
          </div>

{/* 2. ORDERS */}
          <div className="relative">
            <button onClick={() => setSoFilterOpen(!soFilterOpen)} className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md shadow-xs border ${isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600':'border-gray-200 bg-white hover:bg-gray-100'}`}>
              <ShoppingCart weight="bold" /> Orders
              {soTypeApplied && (
                <span className="ml-1 px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-700 dark:bg-indigo-900/50 dark:text-indigo-300 text-[10px] font-bold uppercase tracking-wide">
                  {soTypeIndex.values.find(v => v.key === soTypeApplied)?.label ?? soTypeApplied}
                </span>
              )}
              {selectedSOIds.length > 0 && <span className="ml-1 px-1.5 py-0.5 rounded-full bg-indigo-600 text-white text-xs font-black">{selectedSOIds.length.toLocaleString()}</span>}
            </button>
            {soFilterOpen && (
              <>
                <div className="fixed inset-0 z-30" onClick={() => setSoFilterOpen(false)} />
                <div className={`absolute left-0 top-full mt-1 w-96 max-h-96 overflow-hidden rounded-lg shadow-xl border ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} z-40 flex flex-col`}>
                  <div className={`p-2 border-b ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} flex flex-col gap-2 shrink-0`}>
                    <div className="relative">
                      <MagnifyingGlass className="absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input type="text" placeholder="Search orders..." value={soSearchQuery} onChange={e => setSoSearchQuery(e.target.value)} className={`w-full pl-8 pr-3 py-1.5 text-xs rounded-md border ${isDark?'border-gray-600 bg-gray-800':'border-gray-200 bg-gray-50'} focus:ring-indigo-500`} autoFocus />
                    </div>

                    {/* SO TYPE CHIPS — pick a type to select those orders and hide the rest */}
                    {soTypeIndex.found && soTypeIndex.values.length > 0 && (
                      <div className="flex flex-wrap gap-1">
                        <button
                          onClick={() => { setSoTypeApplied(''); }}
                          className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase tracking-wide border ${!soTypeApplied ? 'bg-indigo-600 text-white border-indigo-600' : isDark?'border-gray-600 text-gray-300 hover:bg-gray-600':'border-gray-200 text-gray-600 hover:bg-gray-100'}`}
                        >
                          All Types
                        </button>
                        {soTypeIndex.values.map(opt => {
                          const count = soTypeIndex.map.get(opt.key)?.length ?? 0;
                          return (
                            <button
                              key={opt.key}
                              onClick={() => {
                                const ids = soTypeIndex.map.get(opt.key) ?? [];
                                setSelectedSOIds(ids);
                                setIsDirty(true);
                                setSoTypeApplied(opt.key);
                              }}
                              className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase tracking-wide border ${soTypeApplied === opt.key ? 'bg-indigo-600 text-white border-indigo-600' : isDark?'border-gray-600 text-gray-300 hover:bg-gray-600':'border-gray-200 text-gray-600 hover:bg-gray-100'}`}
                            >
                              {opt.label} ({count.toLocaleString()})
                            </button>
                          );
                        })}
                      </div>
                    )}

                    <div className="flex justify-between items-center text-[10px] font-black uppercase tracking-widest px-1">
                      <button onClick={() => { setSoTypeApplied(''); handleSOSelectionChange(salesOrderHeaders.map(o => o.id)); }} className="text-indigo-500 hover:text-indigo-600">Select All</button>
                      <button onClick={() => { setSoTypeApplied(''); handleSOSelectionChange([]); }} className="text-rose-500 hover:text-rose-600">Clear</button>
                    </div>
                  </div>
                  <div className="p-2 space-y-1 overflow-y-auto flex-1">
                    {salesOrderHeaders.filter(so => {
                      const q = soSearchQuery.toLowerCase();
                      const matchesSearch = so.name.toLowerCase().includes(q) || so.customer.toLowerCase().includes(q) || (so.customerPO || '').toLowerCase().includes(q);
                      const matchesType = !soTypeApplied || (soTypeIndex.map.get(soTypeApplied) ?? []).includes(so.id);
                      return matchesSearch && matchesType;
                    }).map((so) => (
                      <label key={so.id} className={`flex items-center gap-3 px-2 py-2 rounded cursor-pointer ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${selectedSOIds.includes(so.id) ? (isDark?'bg-gray-600':'bg-gray-100') : ""}`}>
                        <input type="checkbox" checked={selectedSOIds.includes(so.id)} onChange={(e) => handleSOSelectionChange(e.target.checked ? [...selectedSOIds, so.id] : selectedSOIds.filter(id => id !== so.id))} className="rounded text-indigo-600" />
                        <div className="flex-1 min-w-0 flex flex-col">
                          <span className="text-[12px] font-bold truncate">
                            {so.name}{so.customerPO ? ` • PO: ${so.customerPO}` : ''}
                          </span>
                          <span className={`text-[12px] uppercase ${isDark?'text-gray-300':'text-gray-600'} truncate`}>
                            {so.customer} {so.date ? ` • ${formatDate(so.date)}` : ''}
                          </span>
                        </div>
                        <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold uppercase tracking-widest border ${so.status === "Committed" ? "bg-orange-100 text-orange-800 border-orange-200 dark:bg-orange-900/50 dark:text-orange-400 dark:border-orange-800" : "bg-transparent border border-orange-300 text-orange-600"}`}>
                          {so.status === "Committed" ? "Committed" : "Pending"}
                        </span>
                      </label>
                    ))}
                  </div>
                </div>
              </>
            )}
          </div>

          {/* 3. KIT FILTER */}
          {kits.filter(k => (kitDemandMap.get(k.id)?.totalQty ?? 0) > 0).length > 0 && (
            <div className="relative hidden sm:block">
              <button onClick={() => setKitFilterOpen(!kitFilterOpen)} className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md shadow-xs border ${kitFilter ? 'border-purple-400 bg-purple-50 text-purple-700 dark:bg-purple-900/50 dark:border-purple-700 dark:text-purple-400' : isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600':'border-gray-200 bg-white hover:bg-gray-100'}`}>
                <PackageIcon weight="bold" className="text-purple-500" /> Kit
                {kitFilter && <span className="ml-1 px-1.5 py-0.5 rounded-full bg-purple-600 text-white text-xs font-black">1</span>}
              </button>
              {kitFilterOpen && (
                <>
                  <div className="fixed inset-0 z-30" onClick={() => setKitFilterOpen(false)} />
                  <div className={`absolute left-0 top-full mt-1 w-64 max-h-80 overflow-y-auto rounded-lg shadow-lg border ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} z-40`}>
                    <div className={`sticky top-0 px-3 py-2 border-b ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} flex items-center justify-between`}>
                      <span className="text-xs font-semibold">Filter by Kit</span>
                      {kitFilter && <button onClick={() => { setKitFilter(''); setKitFilterOpen(false); }} className="text-xs text-purple-500 font-bold">Clear</button>}
                    </div>
                    <div className="p-2 space-y-1">
                      <button
                        onClick={() => { setKitFilter(''); setKitFilterOpen(false); }}
                        className={`w-full flex items-center gap-2 px-2 py-2 rounded text-left text-xs font-medium ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${!kitFilter ? (isDark?'bg-gray-600 text-gray-100':'bg-gray-100 text-gray-900') : (isDark?'text-gray-300':'text-gray-600')}`}
                      >
                        <PackageIcon weight="regular" className="text-slate-400 shrink-0" />
                        <span>All Kits</span>
                      </button>
                      {kits.filter(k => (kitDemandMap.get(k.id)?.totalQty ?? 0) > 0).map((kit) => (
                        <button
                          key={kit.id}
                          onClick={() => { setKitFilter(kit.id); setKitFilterOpen(false); }}
                          className={`w-full flex items-center gap-2 px-2 py-2 rounded text-left ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${kitFilter === kit.id ? (isDark?'bg-purple-900/50':'bg-purple-50') : ''}`}
                        >
                          <PackageIcon weight={kitFilter === kit.id ? "fill" : "regular"} className="text-purple-500 shrink-0" />
                          <div className="flex-1 min-w-0">
                            <div className="text-xs font-bold truncate">{kit.name}</div>
                            <div className={`text-[9px] uppercase tracking-widest font-bold ${isDark?'text-gray-400':'text-gray-500'}`}>
                              {(kitDemandMap.get(kit.id)?.totalQty ?? 0).toLocaleString()} kits · {kit.productIds.length} products
                            </div>
                          </div>
                          {kitFilter === kit.id && <CheckIcon weight="bold" className="text-purple-500 shrink-0" />}
                        </button>
                      ))}
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {/* 4. CATEGORY */}
          <div className="relative hidden sm:block">
            <button onClick={() => setCategoryFilterOpen(!categoryFilterOpen)} className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md shadow-xs border ${isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600':'border-gray-200 bg-white hover:bg-gray-100'}`}>
              <CubeIcon weight="bold" /> Category
              {categoryFilter.length > 0 && <span className="ml-1 px-1.5 py-0.5 rounded-full bg-orange-600 text-white text-xs font-black">{categoryFilter.length.toLocaleString()}</span>}
            </button>
            {categoryFilterOpen && (
              <>
                <div className="fixed inset-0 z-30" onClick={() => setCategoryFilterOpen(false)} />
                <div className={`absolute left-0 top-full mt-1 w-56 max-h-80 overflow-y-auto rounded-lg shadow-lg border ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} z-40`}>
                  <div className={`sticky top-0 px-3 py-2 border-b ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} flex items-center justify-between`}>
                    <span className="text-xs font-semibold">Filter Category</span>
                    {categoryFilter.length > 0 && <button onClick={() => setCategoryFilter([])} className="text-xs text-orange-500">Clear All</button>}
                  </div>
                  <div className="p-2 space-y-1">
                    {allCategories.map((cat) => (
                      <label key={cat} className={`flex items-center gap-2 px-2 py-1.5 rounded cursor-pointer ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${categoryFilter.includes(cat) ? (isDark?'bg-gray-600':'bg-gray-100') : ""}`}>
                        <input type="checkbox" checked={categoryFilter.includes(cat)} onChange={(e) => setCategoryFilter(p => e.target.checked ? [...p, cat] : p.filter(c => c !== cat))} className="rounded text-orange-500" />
                        <span className="text-xs font-medium">{cat}</span>
                      </label>
                    ))}
                  </div>
                  {categoryFilter.length > 0 && (
                    <div className={`p-2 border-t ${clsStrictBg}`}>
                      <label className="flex items-center gap-2 px-2 py-1.5 rounded cursor-pointer">
                        <input type="checkbox" checked={strictMaterialFilter} onChange={(e) => setStrictMaterialFilter(e.target.checked)} className="rounded text-indigo-600 focus:ring-indigo-500" />
                        <span className={`text-[10px] font-black uppercase tracking-widest ${isDark?'text-indigo-400':'text-indigo-700'}`}>Strict: Hide other materials</span>
                      </label>
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
          
          {/* 5. MATERIALS */}
          <div className="relative hidden sm:block">
            <button onClick={() => setMaterialFilterOpen(!materialFilterOpen)} className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md shadow-xs border ${isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600':'border-gray-200 bg-white hover:bg-gray-100'}`}>
              <PackageIcon weight="bold" /> Materials
              {materialFilter.length > 0 && <span className="ml-1 px-1.5 py-0.5 rounded-full bg-emerald-600 text-white text-xs font-black">{materialFilter.length.toLocaleString()}</span>}
            </button>
            {materialFilterOpen && (
              <>
                <div className="fixed inset-0 z-30" onClick={() => setMaterialFilterOpen(false)} />
                <div className={`absolute left-0 top-full mt-1 w-80 max-h-96 overflow-hidden rounded-lg shadow-xl border ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} z-40 flex flex-col`}>
                  <div className={`p-2 border-b ${isDark?'border-gray-600 bg-gray-700':'border-gray-200 bg-white'} flex flex-col gap-2 shrink-0`}>
                    <div className="relative">
                      <MagnifyingGlass className="absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input type="text" placeholder="Search materials..." value={materialSearchQuery} onChange={e => setMaterialSearchQuery(e.target.value)} className={`w-full pl-8 pr-3 py-1.5 text-xs rounded-md border ${isDark?'border-gray-600 bg-gray-800':'border-gray-200 bg-gray-50'} focus:ring-emerald-500`} autoFocus />
                    </div>
                    <div className="flex justify-between items-center text-[10px] font-black uppercase tracking-widest px-1">
                      <button onClick={() => setMaterialFilter(materials.map(m => m.id))} className="text-emerald-500 hover:text-emerald-600">Select All</button>
                      <button onClick={() => setMaterialFilter([])} className="text-rose-500 hover:text-rose-600">Clear</button>
                    </div>
                  </div>
                  <div className="p-2 space-y-1 overflow-y-auto flex-1">
                    {materials.filter(m => m.name.toLowerCase().includes(materialSearchQuery.toLowerCase())).map((mat) => (
                      <label key={mat.id} className={`flex items-center gap-3 px-2 py-2 rounded cursor-pointer ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'} ${materialFilter.includes(mat.id) ? (isDark?'bg-gray-600':'bg-gray-100') : ""}`}>
                        <input type="checkbox" checked={materialFilter.includes(mat.id)} onChange={(e) => setMaterialFilter(p => e.target.checked ? [...p, mat.id] : p.filter(id => id !== mat.id))} className="rounded text-emerald-600" />
                        <div className="flex-1 min-w-0 flex flex-col">
                          <span className="text-[11px] font-bold truncate">{mat.name}</span>
                          <span className={`text-[9px] uppercase tracking-tighter ${isDark?'text-gray-400':'text-gray-500'} truncate`}>{mat.category} • On Hand: {Math.round(mat.onHand).toLocaleString()}</span>
                        </div>
                      </label>
                    ))}
                  </div>
                </div>
              </>
            )}
          </div>

        </div>

        <div className="ml-auto flex items-center gap-2">
            <Calendar weight="bold" className="w-4 h-4 text-indigo-500" />
            <span className={`text-[10px] font-black uppercase tracking-widest ${isDark?'text-gray-400':'text-gray-500'} hidden sm:block`}>Through</span>
            <input
              type="date"
              value={dateCutoff}
              onChange={(e) => setDateCutoff(e.target.value)}
              className={`px-2 py-1.5 text-xs font-medium rounded-md border ${isDark?'border-gray-600 bg-gray-700 hover:bg-gray-600 text-gray-100':'border-gray-200 bg-white hover:bg-gray-100 text-gray-900'} focus:ring-2 focus:ring-indigo-500 outline-none`}
            />
            {dateCutoff && (
              <button onClick={() => setDateCutoff('')} className={`p-1 rounded ${isDark?'hover:bg-gray-600 text-gray-400':'hover:bg-gray-100 text-gray-500'}`} title="Clear date filter">
                <XIcon weight="bold" className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-auto p-4">
        {visibleSalesOrders.length === 0 ? (
          <div className={`${isDark?'bg-gray-700':'bg-white'} rounded-lg border ${isDark?'border-gray-600':'border-gray-200'} shadow-sm p-8 text-center mt-10`}>
            <PackageIcon weight="duotone" className={`w-12 h-12 mx-auto mb-4 ${isDark?'text-gray-400':'text-gray-500'}`} />
            <h2 className="text-lg font-semibold mb-2">Ready to Plan</h2>
            <p className={`text-sm ${isDark?'text-gray-400':'text-gray-500'} max-w-md mx-auto`}>Select a Plan or pick specific orders from the "Orders" dropdown above to start calculating demand.</p>
          </div>
        ) : (
          <>
            {/* KITS TABLE */}
            {visibleKits.length > 0 && (
              <section className={`${isDark?'bg-gray-700':'bg-white'} rounded-lg border ${isDark?'border-gray-600':'border-gray-200'} shadow-sm mb-4 overflow-hidden`}>
                <div className={`px-4 py-3 border-b ${isDark?'border-gray-600':'border-gray-200'} flex items-center justify-between cursor-pointer hover:bg-black/5 dark:hover:bg-white/5 transition-colors`} onClick={() => setShowKits(!showKits)}>
                  <div className="flex items-center gap-2">
                    <PackageIcon weight="duotone" className="w-5 h-5 text-purple-500" />
                    <span className="font-semibold text-sm">Kits Demand</span>
                    <span className={`text-xs px-2 py-0.5 rounded-full font-bold ${clsKitsBadge}`}>{visibleKits.length}</span>
                  </div>
                  <CaretRightIcon className={`w-4 h-4 transition-transform ${showKits ? 'rotate-90' : ''}`} />
                </div>
                {showKits && (
                <div className="overflow-x-auto relative">
                  <table className="w-full text-xs">
                    <thead className={isDark?'bg-gray-600':'bg-gray-100'}>
                      <tr>
                        <th className={`px-3 py-2 text-left min-w-[220px] max-w-[220px] w-[220px] sticky left-0 z-20 ${isDark?'bg-gray-600':'bg-gray-100'} border-r ${isDark?'border-gray-600':'border-gray-200'}`}>Kit</th>
                        <th className={`px-3 py-2 text-center min-w-[110px] max-w-[110px] w-[110px] sticky left-[220px] z-20 ${isDark?'bg-gray-600':'bg-gray-100'} border-r ${isDark?'border-gray-600':'border-gray-200'} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>Total Demand</th>
                        {activeDates.map((date) => (
                          <th key={date} className={`px-2 py-2 text-center min-w-[100px] border-l ${isDark?'border-gray-600':'border-gray-200'}`}>
                            <span className="font-bold">{formatDate(date)}</span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleKits.map((kit) => {
                        const demand = kitDemandMap.get(kit.id) ?? { totalQty: 0, inMoQty: 0, moNames: new Set<string>(), byDate: new Map(), inMoByDate: new Map<string, number>(), moNamesByDate: new Map<string, Set<string>>() };
                        return (
                          <tr key={kit.id} className={`border-b ${isDark?'border-gray-600':'border-gray-200'} ${isDark?'hover:bg-gray-600':'hover:bg-gray-50'} group`}>
                            <td className={`px-3 py-2 font-medium min-w-[220px] max-w-[220px] w-[220px] sticky left-0 z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} transition-colors border-r ${isDark?'border-gray-600':'border-gray-200'}`}>
                              <div className="truncate" title={kit.name}>{kit.name}</div>
                            </td>
                            <td className={`px-3 py-2 text-center tabular-nums min-w-[110px] max-w-[110px] w-[110px] sticky left-[220px] z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} transition-colors border-r ${isDark?'border-gray-600':'border-gray-200'} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>
<span 
                                className="inline-flex items-center px-2 py-1 rounded bg-purple-100 text-purple-800 border border-purple-200 dark:bg-purple-900/50 dark:text-purple-400 dark:border-purple-800 font-black cursor-pointer hover:ring-2 hover:ring-purple-400 transition-all text-sm"
                                onClick={(e) => handlePillClick(e, renderKitDemandPopup(kit, demand.totalQty))}
                              >
                                {Math.round(demand.totalQty).toLocaleString()}
                              </span>                            </td>
                            {activeDates.map((date) => {
                              const qty = demand.byDate.get(date) ?? 0;
                              const dateInMo = demand.inMoByDate.get(date) ?? 0;
                              const dateFullyInMO = qty > 0 && dateInMo >= qty;
                              const dateMoNames = Array.from(demand.moNamesByDate.get(date) ?? []).join(', ');
                              return (
                                <td key={date} className={`px-2 py-2 border-l ${isDark?'border-gray-600':'border-gray-200'} text-center`}>
                                  {qty > 0 && (
                                    <div className="flex flex-col items-center gap-0.5">
                                      <span
                                        className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-black cursor-pointer transition-all border ${
                                          dateFullyInMO
                                            ? 'bg-indigo-100 text-indigo-800 border-indigo-200 dark:bg-indigo-900/50 dark:text-indigo-400 dark:border-indigo-800 hover:ring-2 hover:ring-indigo-400'
                                            : 'bg-transparent border-purple-300 text-purple-700 dark:border-purple-700 dark:text-purple-400 hover:bg-purple-50 dark:hover:bg-purple-900/30'
                                        }`}
                                        onClick={(e) => handlePillClick(e, renderKitDemandPopup(kit, qty, date))}
                                      >
                                        {Math.round(qty).toLocaleString()}
                                      </span>
                                      {dateMoNames && (
                                        <span
                                          className="text-[9px] font-black uppercase tracking-tighter text-indigo-600 dark:text-indigo-400 truncate max-w-[92px]"
                                          title={dateMoNames}
                                        >
                                          {dateFullyInMO ? '' : 'partial · '}{dateMoNames}
                                        </span>
                                      )}
                                    </div>
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                )}
              </section>
            )}

            {/* FG TABLE */}
            <section className={`${isDark?'bg-gray-700':'bg-white'} rounded-lg border ${isDark?'border-gray-600':'border-gray-200'} shadow-sm mb-4 overflow-hidden`}>
              <div className={`px-4 py-3 border-b ${isDark?'border-gray-600':'border-gray-200'} flex items-center justify-between cursor-pointer hover:bg-black/5 dark:hover:bg-white/5 transition-colors`} onClick={() => setShowFG(!showFG)}>
                <div className="flex items-center gap-2">
                  <CubeIcon weight="duotone" className="w-5 h-5 text-orange-500" />
                  <span className="font-semibold text-sm">Finished Goods Demand</span>
                </div>
                <CaretRightIcon className={`w-4 h-4 transition-transform ${showFG ? 'rotate-90' : ''}`} />
              </div>
              
              {showFG && (
                <div className="overflow-x-auto relative">
                  <table className="w-full text-xs">
                    <thead className={isDark?'bg-gray-600':'bg-gray-100'}>
                      <tr>
                        <th className={`px-3 py-2 text-left min-w-[200px] max-w-[200px] w-[200px] sticky left-0 z-20 ${isDark?'bg-gray-600':'bg-gray-100'} border-r ${isDark?'border-gray-600':'border-gray-200'}`}>Product</th>
                        <th className={`px-3 py-2 text-center min-w-[100px] max-w-[100px] w-[100px] sticky left-[200px] z-20 ${isDark?'bg-gray-600':'bg-gray-100'} border-r ${isDark?'border-gray-600':'border-gray-200'} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>Committed</th>
                        <th className={`px-3 py-2 text-center min-w-[100px] border-r ${borderColor}`}>Total Demand</th>
                        {activeDates.map((date) => (
                          <th key={date} className={`px-2 py-2 text-center min-w-[100px] border-l ${isDark?'border-gray-600':'border-gray-200'}`}>
                            <span className="font-bold text-gray-900 dark:text-gray-100">{formatDate(date)}</span>
                          </th>
                        ))}
                        <th className="px-2 py-2 w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {productIds.map((pid) => {
                        const product = products.find((p) => p.id === pid);
                        if (!product) return null;
                        let totalCommitted = 0;
                        const cSOs: SalesOrderHeader[] = [];
                        committedSOsSandbox.forEach((so) => {
                          const lis = getLineItemsForOrder(so.id).filter(li => li.productId === pid);
                          if (lis.length > 0) {
                            totalCommitted += lis.reduce((s, li) => s + li.qty, 0);
                            cSOs.push(so);
                          }
                        });
                        
                        return (
                          <tr key={pid} className={`border-b ${isDark?'border-gray-600':'border-gray-200'} ${isDark?'hover:bg-gray-600':'hover:bg-gray-50'} group`}>
                            <td className={`px-3 py-2 font-medium min-w-[200px] max-w-[200px] w-[200px] sticky left-0 z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} transition-colors border-r ${isDark?'border-gray-600':'border-gray-200'}`}>
                              <div className="truncate" title={product.name}>{product.name}</div>
                            </td>
                            <td className={`px-3 py-2 text-center tabular-nums min-w-[100px] max-w-[100px] w-[100px] sticky left-[200px] z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} transition-colors border-r ${isDark?'border-gray-600':'border-gray-200'} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>
                              {totalCommitted > 0 && (
                                <span
                                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-orange-100 text-orange-800 border border-orange-200 dark:bg-orange-900/50 dark:text-orange-400 dark:border-orange-800 text-xs font-black cursor-pointer"
                                  onClick={(e) => handlePillClick(e, 
                                    <div className="text-xs max-w-[320px] max-h-64 overflow-y-auto">
                                      <div className={`font-semibold mb-2 sticky top-0 ${isDark?'bg-gray-700':'bg-white'} pb-1 z-10`}>Committed Orders ({Math.round(totalCommitted).toLocaleString()})</div>
                                      {cSOs.map((so) => (
                                        <button onClick={(e) => expandSO(e, so.id)} key={so.id} className={`w-full text-left mb-1 p-2 rounded border ${borderColor} ${isDark ? 'hover:bg-gray-600' : 'hover:bg-gray-50'}`}>
                                          <div className="flex justify-between gap-4">
                                            <span className="font-bold text-slate-900 dark:text-slate-100">{so.customer} {so.customerPO ? `(PO: ${so.customerPO})` : ''}: {Math.round(getLineItemsForOrder(so.id).filter(li => li.productId === pid).reduce((s, li) => s + li.qty, 0)).toLocaleString()}</span>
                                            <span className="text-slate-500">{formatDate(so.date)}</span>
                                          </div>
                                        </button>
                                      ))}
                                    </div>
                                  )}
                                >
                                  <WarningIcon weight="bold" /> {Math.round(totalCommitted).toLocaleString()}
                                </span>
                              )}
                            </td>
                            <td className={`px-3 py-2 text-center tabular-nums border-r ${borderColor}`}>
                              {visibleSalesOrders.reduce((sum, so) => sum + getLineItemsForOrder(so.id).filter(li => li.productId === pid).reduce((s, li) => s + li.qty, 0), 0) > 0 && (
                                <span className="text-sm font-black tabular-nums text-slate-700 dark:text-slate-200">
                                  {Math.round(visibleSalesOrders.reduce((sum, so) => sum + getLineItemsForOrder(so.id).filter(li => li.productId === pid).reduce((s, li) => s + li.qty, 0), 0)).toLocaleString()}
                                </span>
                              )}
                            </td>
                            {activeDates.map((date) => {
                              const bucket = computeFGBuckets(pid, date);
                              const hasActivity = bucket.committedReady > 0 || bucket.committedShort > 0 || bucket.atpReady > 0 || bucket.late > 0 || bucket.short > 0;
                              if (!hasActivity) return <td key={date} className={`px-2 py-2 border-l ${borderColor}`}></td>;
                              
                              const renderTooltip = (title: string, qty: number, sos: SalesOrderHeader[], colorText: string, colorBg: string) => (
                                <div className="text-xs w-[380px] max-h-[400px] overflow-y-auto pr-1 flex flex-col gap-3">
                                  <div className={`font-black uppercase tracking-widest px-3 py-2 rounded-lg sticky top-0 z-10 border ${colorText} ${colorBg} shadow-sm backdrop-blur-md bg-opacity-95`}>
                                    {title} &bull; {Math.round(qty).toLocaleString()} UNITS
                                  </div>
                                  <div className="flex flex-col gap-3 pb-2">
                                    {sos.map((so) => {
                                      const lineItems = getLineItemsForOrder(so.id);
                                      const soQty = lineItems.filter(li => li.productId === pid).reduce((sum, li) => sum + li.qty, 0);
                                      const bomEntries = bom.filter(b => b.productId === pid);
                                      return (
                                        <button onClick={(e) => expandSO(e, so.id)} key={so.id} className={`w-full text-left p-3 rounded-xl border ${borderColor} ${isDark ? 'bg-gray-800 hover:bg-gray-700' : 'bg-white hover:bg-gray-50'} shadow-sm block transition-colors`}>
                                          <div className="flex justify-between items-start mb-2">
                                            <div className="flex flex-col">
                                              <span className="font-bold text-sm uppercase text-slate-900 dark:text-slate-100">{so.customer}</span>
                                              <span className={`text-[10px] ${textSecondary} uppercase tracking-wider`}>{so.name} {so.customerPO ? ` • PO: ${so.customerPO}` : ''} • {so.soStatus || 'Pending'}</span>
                                            </div>
                                            <span className="font-mono font-black text-blue-600 bg-blue-50 dark:bg-blue-900/30 dark:text-blue-400 px-2 py-0.5 rounded border border-blue-100 dark:border-blue-800">
                                              {Math.round(soQty).toLocaleString()} FG
                                            </span>
                                          </div>
                                          {bomEntries.length > 0 && (
                                            <div className={`mt-3 flex flex-col rounded-lg border ${borderColor} bg-slate-50 dark:bg-gray-900/50 overflow-hidden`}>
                                              <div className={`px-3 py-1.5 text-[9px] font-black uppercase tracking-widest ${textSecondary} bg-slate-100 dark:bg-gray-800 border-b ${borderColor}`}>BOM Requirements</div>
                                              <div className="flex flex-col px-3 py-1">
                                                {bomEntries.map((bomEntry, idx) => {
                                                  const mat = materials.find(m => m.id === bomEntry.materialId);
                                                  const qtyNeeded = bomEntry.qtyPer * soQty;
                                                  const wf = allSandboxWaterfalls.get(bomEntry.materialId);
                                                  const dateData = wf?.get(date);
                                                  const balOnDate = dateData?.balance ?? (availableNowSandboxByMaterial.get(bomEntry.materialId) ?? 0);
                                                  const shortQty = balOnDate < 0 ? Math.min(qtyNeeded, Math.abs(balOnDate)) : 0;
                                                  let statusLabel = '';
                                                  let statusColor = '';
                                                  if (shortQty === 0) {
                                                    let incomingUpToDate = 0;
                                                    if (wf) { for (const [d, val] of wf.entries()) { if (d <= date) incomingUpToDate += val.incoming; } }
                                                    const balWithoutPOs = balOnDate - incomingUpToDate;
                                                    if (balWithoutPOs >= 0 || incomingUpToDate === 0) { statusLabel = 'IN STOCK'; statusColor = 'bg-slate-200 text-slate-700 border-slate-300 dark:bg-gray-700 dark:text-gray-300 dark:border-gray-600'; }
                                                    else { statusLabel = 'ON TIME PO'; statusColor = 'bg-indigo-100 text-indigo-700 border-indigo-200 dark:bg-indigo-900/50 dark:text-indigo-400 dark:border-indigo-800'; }
                                                  } else {
                                                    let futureIncoming = 0;
                                                    if (wf) { for (const [d, val] of wf.entries()) { if (d > date) futureIncoming += val.incoming; } }
                                                    if (futureIncoming >= shortQty) { statusLabel = 'LATE PO'; statusColor = 'bg-purple-100 text-purple-700 border-purple-200 dark:bg-purple-900/50 dark:text-purple-400 dark:border-purple-800'; }
                                                    else { statusLabel = `SHORT (${Math.round(shortQty).toLocaleString()})`; statusColor = 'bg-rose-100 text-rose-700 border-rose-200 dark:bg-rose-900/50 dark:text-rose-400 dark:border-rose-800'; }
                                                  }
                                                  return (
                                                    <div key={idx} className={`flex justify-between items-center gap-4 py-2 border-b ${borderColor} last:border-0`}>
                                                      <div className="flex flex-col min-w-0 flex-1">
                                                        <span className="truncate text-[11px] font-bold text-slate-900 dark:text-slate-100" title={mat?.name}>{mat?.name || 'Unknown'}</span>
                                                        <span className={`text-[9px] uppercase font-bold tracking-widest ${textSecondary}`}>{Math.round(qtyNeeded).toLocaleString()} REQ</span>
                                                      </div>
                                                      <span className={`shrink-0 px-2 py-1 rounded-md text-[9px] font-black uppercase tracking-widest border ${statusColor}`}>{statusLabel}</span>
                                                    </div>
                                                  );
                                                })}
                                              </div>
                                            </div>
                                          )}
                                        </button>
                                      );
                                    })}
                                  </div>
                                </div>
                              );

                              return (
                                <td key={date} className={`px-2 py-2 border-l ${borderColor}`}>
                                  <div className="flex flex-wrap gap-1 justify-center">
                                    {bucket.committedReady > 0 && (
                                      <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 border border-blue-200 dark:bg-blue-900/50 dark:text-blue-400 dark:border-blue-800 cursor-pointer text-xs font-black transition-all hover:ring-2 hover:ring-blue-400"
                                        onClick={(e) => handlePillClick(e, renderTooltip("Committed & Ready", bucket.committedReady, bucket.committedReadySOs, "text-blue-800 dark:text-blue-300", "bg-blue-100 border-blue-200 dark:bg-blue-900/80 dark:border-blue-700"))}>
                                        <CheckIcon /> {Math.round(bucket.committedReady).toLocaleString()}
                                      </span>
                                    )}
                                    {bucket.committedShort > 0 && (
                                      <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded border border-rose-300 dark:border-rose-800 bg-rose-100 dark:bg-rose-900/50 text-rose-800 dark:text-rose-400 cursor-pointer font-black text-xs transition-all hover:ring-2 hover:ring-rose-400"
                                        onClick={(e) => handlePillClick(e, renderTooltip("Committed BUT Short", bucket.committedShort, bucket.committedShortSOs, "text-rose-800 dark:text-rose-300", "bg-rose-100 border-rose-200 dark:bg-rose-900/80 dark:border-rose-700"))}>
                                        <WarningIcon /> {Math.round(bucket.committedShort).toLocaleString()}
                                      </span>
                                    )}
                                    {bucket.atpReady > 0 && (
                                      <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded bg-transparent border border-emerald-300 text-emerald-700 dark:text-emerald-400 dark:border-emerald-700 cursor-pointer text-xs font-black transition-all hover:bg-emerald-50 dark:hover:bg-emerald-900/30"
                                        onClick={(e) => handlePillClick(e, renderTooltip("ATP Ready", bucket.atpReady, bucket.atpReadySOs, "text-emerald-800 dark:text-emerald-300", "bg-transparent border-emerald-300 dark:border-emerald-700"))}>
                                        {Math.round(bucket.atpReady).toLocaleString()}
                                      </span>
                                    )}
                                    {bucket.late > 0 && (
                                      <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded bg-transparent border border-orange-300 text-orange-700 dark:text-orange-400 dark:border-orange-700 cursor-pointer text-xs font-black transition-all hover:bg-orange-50 dark:hover:bg-orange-900/30"
                                        onClick={(e) => handlePillClick(e, renderTooltip("Late", bucket.late, bucket.lateSOs, "text-orange-800 dark:text-orange-300", "bg-transparent border-orange-300 dark:border-orange-700"))}>
                                        {Math.round(bucket.late).toLocaleString()}
                                      </span>
                                    )}
                                    {bucket.short > 0 && (
                                      <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded bg-transparent border border-rose-300 text-rose-700 dark:text-rose-400 dark:border-rose-700 cursor-pointer text-xs font-black transition-all hover:bg-rose-50 dark:hover:bg-rose-900/30"
                                        onClick={(e) => handlePillClick(e, renderTooltip("Short", bucket.short, bucket.shortSOs, "text-rose-800 dark:text-rose-300", "bg-transparent border-rose-300 dark:border-rose-700"))}>
                                        {Math.round(bucket.short).toLocaleString()}
                                      </span>
                                    )}
                                  </div>
                                </td>
                              );
                            })}
                            <td className="px-2 py-2 text-center">
                              <button onClick={() => openFlyout("product", pid)} className={`p-1 rounded ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'}`}><CaretRightIcon /></button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>

            {/* RM TABLE */}
            <section className={`${bgCard} rounded-lg border ${borderColor} shadow-sm overflow-hidden`}>
              <div className={`px-4 py-3 border-b ${borderColor} flex items-center justify-between cursor-pointer hover:bg-black/5 dark:hover:bg-white/5 transition-colors`} onClick={() => setShowRM(!showRM)}>
                <div className="flex items-center gap-2">
                  <PackageIcon weight="duotone" className="w-5 h-5 text-emerald-500" />
                  <span className="font-semibold text-sm">Raw Materials / BOM Plan</span>
                  <button onClick={(e) => { e.stopPropagation(); setIncludeOtherMoNeed(v => !v); }}
  className={`ml-3 px-2 py-1 rounded-lg text-[11px] font-bold border transition-colors ${includeOtherMoNeed
    ? 'bg-indigo-100 text-indigo-700 border-indigo-300 dark:bg-indigo-900/40 dark:text-indigo-300 dark:border-indigo-700'
    : (isDark ? 'bg-gray-700 text-gray-300 border-gray-600' : 'bg-gray-100 text-gray-600 border-gray-300')}`}
  title="Include unpicked MOs not linked to the selected orders. Affects pills and Bal — All always includes everything.">
  {includeOtherMoNeed ? 'Other MOs: shown' : 'Other MOs: hidden'}
</button>
                </div>
                <CaretRightIcon className={`w-4 h-4 transition-transform ${showRM ? 'rotate-90' : ''}`} />
              </div>
              
              {showRM && (
                <div className="overflow-x-auto relative">
                  <table className="w-full text-xs">
                    <thead className={theadBg}>
                      <tr>
                        <th className={`px-3 py-2 text-left min-w-[200px] max-w-[200px] w-[200px] sticky left-0 z-20 ${theadBg} border-r ${borderColor}`}>Material</th>
                        <th className={`px-3 py-2 text-center min-w-[120px] max-w-[120px] w-[120px] sticky left-[200px] z-20 ${theadBg} border-r ${borderColor}`}>Ordered</th>
                        <th className={`px-3 py-2 text-center min-w-[100px] max-w-[100px] w-[100px] sticky left-[320px] z-20 ${theadBg} border-r ${borderColor}`}>On Hand / Com.</th>
                        <th className={`px-3 py-2 text-center min-w-[80px] max-w-[80px] w-[80px] sticky left-[420px] z-20 ${theadBg} border-r ${borderColor} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>Avail Now</th>
                        {activeDates.map((date) => (
                          <th key={date} className={`px-2 py-2 text-center min-w-[100px] border-l ${borderColor}`}>
                            <span className="font-bold text-gray-900 dark:text-gray-100">{formatDate(date)}</span>
                          </th>
                        ))}
                        <th className="px-2 py-2 w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {materialIds.map((matId) => {
                        const material = materials.find((m) => m.id === matId);
                        if (!material) return null;
                        
                        const sandboxWaterfall = allSandboxWaterfalls.get(matId);
                        const rwWaterfall = allRealWorldWaterfalls.get(matId);
                        
                        const committedQty = totalCommittedSandboxByMaterial.get(matId) ?? 0;
                        const availNow = availableNowSandboxByMaterial.get(matId) ?? 0;

                        let totalReady = 0;
                        let totalOnOrder = 0;
                        if (material.orderedStr) {
                          const regex = /([\d,.]+)\s*(ready|on order)/gi;
                          let match;
                          while ((match = regex.exec(material.orderedStr)) !== null) {
                            const numStr = match[1].replace(/[^\d.]/g, '');
                            const qty = parseFloat(numStr) || 0;
                            if (match[2].toLowerCase() === 'ready') totalReady += qty;
                            else if (match[2].toLowerCase() === 'on order') totalOnOrder += qty;
                          }
                        }

                        return (
                          <tr key={matId} className={`border-b ${borderColor} ${isDark?'hover:bg-gray-600':'hover:bg-gray-50'} group`}>
                            <td className={`px-3 py-2 font-medium min-w-[200px] max-w-[200px] w-[200px] sticky left-0 z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} transition-colors border-r ${borderColor}`}>
                              <div className="truncate" title={material.name}>{material.name}</div>
                            </td>
                            <td className={`px-2 py-2 text-center min-w-[120px] max-w-[120px] w-[120px] sticky left-[200px] z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} border-r ${borderColor}`}>
                              <div className="flex flex-col gap-0.5 items-center justify-center w-full py-0.5">
                                {totalOnOrder > 0 && (
                                  <div className="flex items-center justify-between w-[95%] px-1">
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-transparent border border-indigo-300 text-indigo-700 dark:border-indigo-700 dark:text-indigo-400">ON ORD</span>
                                    <span className="text-xs font-black tabular-nums text-slate-700 dark:text-slate-200">{Math.round(totalOnOrder).toLocaleString()}</span>
                                  </div>
                                )}
                                {totalReady > 0 && (
                                  <div className="flex items-center justify-between w-[95%] px-1">
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-blue-100 text-blue-800 border border-blue-200 dark:bg-blue-900/50 dark:text-blue-400 dark:border-blue-800">READY</span>
                                    <span className="text-xs font-black tabular-nums text-slate-700 dark:text-slate-200">{Math.round(totalReady).toLocaleString()}</span>
                                  </div>
                                )}
                                {(totalReady === 0 && totalOnOrder === 0) && (
                                  <span className={`text-[11px] ${isDark?'text-gray-500':'text-gray-300'}`}>—</span>
                                )}
                              </div>
                            </td>
                            <td className={`px-3 py-2 text-center min-w-[100px] max-w-[100px] w-[100px] sticky left-[320px] z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} border-r ${borderColor}`}>
                              <div className="flex flex-col items-center justify-center gap-0.5 py-0.5">
                                <span className="text-xs font-black tabular-nums text-slate-700 dark:text-slate-200 leading-none">{Math.round(material.onHand).toLocaleString()}</span>
                                {committedQty > 0 ? (
                                  <span className="inline-flex items-center gap-0.5 px-1 py-0 rounded bg-orange-100 text-orange-800 border border-orange-200 dark:bg-orange-900/50 dark:text-orange-400 dark:border-orange-800 text-[9px] font-bold tabular-nums mt-0.5">
                                    <WarningIcon weight="bold" size={10} /> {Math.round(committedQty).toLocaleString()}
                                  </span>
                                ) : (
                                  <span className={`text-[10px] ${isDark?'text-gray-500':'text-gray-300'}`}>-</span>
                                )}
                              </div>
                            </td>
                            <td className={`px-3 py-2 text-center min-w-[80px] max-w-[80px] w-[80px] sticky left-[420px] z-10 ${isDark?'bg-gray-700 group-hover:bg-gray-600':'bg-white group-hover:bg-gray-50'} border-r ${borderColor} shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]`}>
                               {availNow >= 0 ? (
                                 <span className="text-sm font-black text-emerald-600 dark:text-emerald-400 tabular-nums leading-none">{Math.round(availNow).toLocaleString()}</span>
                               ) : (
                                 <span className="inline-block px-1.5 py-0.5 rounded bg-rose-50 border border-rose-200 text-rose-600 dark:bg-rose-900/30 dark:border-rose-800 dark:text-rose-400 text-sm font-black tabular-nums leading-none">{Math.round(availNow).toLocaleString()}</span>
                               )}
                            </td>
                            {activeDates.map((date) => {
                              const data = sandboxWaterfall?.get(date);
                              const rwData = rwWaterfall?.get(date);
                              const soNeed = unreleasedSODemandByMaterialDate.get(matId)?.get(date) ?? 0;
                              const moNeed = effectiveMoNeedByMaterialDate.get(matId)?.get(date) ?? 0;
                              const totNeed = soNeed + moNeed;
                              const needLabel = soNeed > 0 && moNeed > 0
                                ? `SO ${Math.round(soNeed).toLocaleString()} + MO ${Math.round(moNeed).toLocaleString()}`
                                : soNeed > 0 ? `SO ${Math.round(soNeed).toLocaleString()}` : `MO ${Math.round(moNeed).toLocaleString()}`;
                              const moRequiredHere = (includeOtherMoNeed ? moMapsForDisplay.requiredAll : moMapsForDisplay.requiredSelected).get(matId)?.get(date) ?? 0;
                              const fullyPickedHere = totNeed <= 0 && moRequiredHere > 0;
                              if (!data) return <td key={date} className={`px-2 py-2 border-l ${borderColor}`}></td>;
                              if (data.incoming === 0 && data.timePhasedDemand === 0 && totNeed <= 0 && !fullyPickedHere) return <td key={date} className={`px-2 py-2 border-l ${borderColor}`}></td>;
                              
                              const isSandboxSafe = data.balance >= 0;
                              const isRealWorldSafe = rwData && rwData.balance >= 0;
                              
                              let demandColorClass = "";
                              if (isRealWorldSafe) {
                                  demandColorClass = "border-emerald-300 text-emerald-700 dark:border-emerald-700 dark:text-emerald-400 hover:bg-emerald-50 dark:hover:bg-emerald-900/30";
                              } else if (isSandboxSafe) {
                                  demandColorClass = "border-yellow-400 text-yellow-700 dark:border-yellow-600 dark:text-yellow-400 hover:bg-yellow-50 dark:hover:bg-yellow-900/30 bg-yellow-50/50 dark:bg-yellow-900/20";
                              } else {
                                  demandColorClass = "border-rose-300 text-rose-700 dark:border-rose-700 dark:text-rose-400 hover:bg-rose-50 dark:hover:bg-rose-900/30";
                              }

                              return (
                                <td key={date} className={`px-2 py-2 border-l ${borderColor}`}>
                                  <div className="flex flex-col items-center gap-1">
                                    {data.incoming > 0 && (
                                      <span 
                                        className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 border border-blue-200 dark:bg-blue-900/30 dark:text-blue-400 dark:border-blue-800 text-[10px] font-black cursor-pointer hover:bg-blue-100 dark:hover:bg-blue-900/50 transition-colors"
                                        onClick={(e) => {
                                          const pos = purchaseOrders.filter((po) => po.materialId === matId && po.date === date);
                                          const vps = vPos.filter((vpo) => vpo.materialId === matId && vpo.date === date);
                                          handlePillClick(e, (
                                            <div className="text-xs w-full pr-1 flex flex-col gap-3">
                                              <div className="font-black uppercase tracking-widest px-3 py-2 rounded-lg sticky top-0 z-10 border text-blue-800 dark:text-blue-300 bg-blue-100 border-blue-200 dark:bg-blue-900/80 dark:border-blue-700 shadow-sm backdrop-blur-md bg-opacity-95">
                                                INCOMING SUPPLY &bull; +{Math.round(data.incoming).toLocaleString()} UNITS
                                              </div>
                                              <div className="flex flex-col gap-3 pb-2">
                                                {pos.map((po) => (
                                                  <button onClick={(e) => expandPO(e, po.headerId)} key={po.id} className={`w-full text-left p-3 rounded-xl border ${borderColor} ${isDark ? 'bg-gray-800 hover:bg-gray-700' : 'bg-white hover:bg-gray-50'} shadow-sm block transition-colors`}>
                                                    <div className="flex justify-between items-center">
                                                      <div className="flex flex-col">
                                                        <span className="font-bold text-sm uppercase text-slate-900 dark:text-slate-100">{po.vendor}</span>
                                                        <span className={`text-[10px] ${textSecondary} uppercase tracking-wider`}>{po.isShipment ? `Shipment: ${po.shipmentName || 'Linked'} (PO ${po.name || po.id})` : `PO ${po.name || po.id} (Unscheduled)`}</span>
                                                      </div>
                                                      <span className="font-mono font-black text-blue-600 bg-blue-50 dark:bg-blue-900/30 dark:text-blue-400 px-2 py-0.5 rounded border border-blue-100 dark:border-blue-800">
                                                        +{Math.round(po.qty).toLocaleString()}
                                                      </span>
                                                    </div>
                                                  </button>
                                                ))}
                                                {vps.map((vpo) => (
                                                  <div key={vpo.id} className="p-3 rounded-xl border border-purple-200 dark:border-purple-800 bg-purple-50 dark:bg-purple-900/20 shadow-sm flex justify-between items-center">
                                                    <div className="flex flex-col">
                                                      <span className="font-bold text-sm text-purple-700 dark:text-purple-400">SIMULATED PO</span>
                                                      <span className={`text-[10px] ${textSecondary} uppercase tracking-wider`}>{vpo.vendor}</span>
                                                    </div>
                                                    <span className="font-mono font-black text-purple-600 bg-purple-100 dark:bg-purple-900/50 dark:text-purple-300 px-2 py-0.5 rounded border border-purple-200 dark:border-purple-800">
                                                      +{Math.round(vpo.qty).toLocaleString()}
                                                    </span>
                                                  </div>
                                                ))}
                                              </div>
                                            </div>
                                          ));
                                        }}
                                      >
                                        +{Math.round(data.incoming).toLocaleString()}
                                      </span>
                                    )}
                                    {totNeed > 0 && (
                                      <span 
                                        className={`px-1.5 py-0.5 rounded bg-transparent border ${demandColorClass} text-[10px] font-black cursor-pointer transition-colors`}
                                        onClick={(e) => {
                                          const sos = timePhasedSOsSandbox.filter((so) => so.date === date);
                                          handlePillClick(e, (
                                            <div className="text-xs w-full pr-1 flex flex-col gap-3">
                                              <div className={`font-black uppercase tracking-widest px-3 py-2 rounded-lg sticky top-0 z-10 border text-orange-800 dark:text-orange-300 bg-orange-100 border-orange-200 dark:bg-orange-900/80 dark:border-orange-700 shadow-sm backdrop-blur-md bg-opacity-95`}>
                                                Still Needed &bull; -{Math.round(totNeed).toLocaleString()} UNITS
                                              </div>
                                              <div className="flex flex-col gap-3 pb-2">
                                                {sos.map((so) => {
                                                  const demand = getMaterialDemandFromOrder(so).get(matId) ?? 0;
                                                  if (demand === 0) return null;
                                                  return (
                                                    <button onClick={(e) => expandSO(e, so.id)} key={so.id} className={`w-full text-left p-3 rounded-xl border ${borderColor} ${isDark ? 'bg-gray-800 hover:bg-gray-700' : 'bg-white hover:bg-gray-50'} shadow-sm block transition-colors`}>
                                                      <div className="flex justify-between items-center">
                                                        <div className="flex flex-col">
                                                          <span className="font-bold text-sm text-slate-900 dark:text-slate-100">{so.customer}</span>
                                                          <span className={`text-[12px] ${textSecondary} uppercase tracking-wider`}>{so.name} {so.customerPO ? ` • PO: ${so.customerPO}` : ''}</span>
                                                        </div>
                                                        <span className="font-mono font-black text-orange-600 bg-transparent border border-orange-300 dark:border-orange-700 dark:text-orange-400 px-2 py-0.5 rounded">
                                                          -{Math.round(demand).toLocaleString()}
                                                        </span>
                                                      </div>
                                                    </button>
                                                  );
                                                })}
                                                {(stillToPickDetailByMaterialDate.get(matId)?.get(date) ?? []).filter(mi => includeOtherMoNeed || mi.isSelected).map((moItem) => (
                                                  <button
                                                    onClick={(ev) => { ev.preventDefault(); ev.stopPropagation(); const rec = moRecords.find(r => r.id === moItem.moRecId); if (rec) expandRecord(rec); }}
                                                    key={moItem.moRecId}
                                                    className={`w-full text-left p-3 rounded-xl border border-indigo-200 dark:border-indigo-800 ${isDark ? 'bg-indigo-900/20 hover:bg-indigo-900/40' : 'bg-indigo-50 hover:bg-indigo-100'} shadow-sm block transition-colors`}
                                                  >
                                                    <div className="flex justify-between items-center">
                                                      <div className="flex flex-col">
                                                        <span className="font-bold text-sm text-indigo-700 dark:text-indigo-300">{moItem.moName}</span>
                                                        <span className={`text-[12px] ${textSecondary} uppercase tracking-wider`}>{moItem.soLabel ? `for ${moItem.soLabel}` : 'no SO linked'} • unpicked</span>
                                                      </div>
                                                      <span className="font-mono font-black text-indigo-600 bg-transparent border border-indigo-300 dark:border-indigo-700 dark:text-indigo-400 px-2 py-0.5 rounded">
                                                        -{Math.round(moItem.qty).toLocaleString()}
                                                      </span>
                                                    </div>
                                                  </button>
                                                ))}
                                              </div>
                                            </div>
                                          ));
                                        }}
                                      >
                                        -{Math.round(totNeed).toLocaleString()}
                                      </span>
                                    )}
                                    {moNeed > 0 && totNeed > moNeed && (
                                      <span className={`text-[8px] uppercase font-bold tracking-tighter ${textSecondary} tabular-nums`}>MO {Math.round(moNeed).toLocaleString()} unpicked</span>
                                    )}
                                    {fullyPickedHere && (
                                      <div className="flex flex-col items-center gap-0.5">
                                        <span className={`text-xs font-black tabular-nums ${textSecondary}`}>{Math.round(moRequiredHere).toLocaleString()}</span>
                                        <span className={`text-[8px] uppercase font-bold tracking-tighter ${textSecondary}`}>picked</span>
                                      </div>
                                    )}
                                    {(() => {
                                      const shelfBal = selectedShelfBalByMaterialDate.get(matId)?.get(date);
                                      const allBal = rwData?.balance;
                                      return (
                                        <div className="flex flex-col items-center mt-1">
                                          <div className={`text-[10px] tabular-nums font-medium ${(shelfBal ?? 0) < 0 ? "text-rose-600 font-bold" : textSecondary}`}>Bal: {shelfBal !== undefined ? Math.round(shelfBal).toLocaleString() : '—'}</div>
                                          <div className={`text-[10px] tabular-nums font-medium ${(allBal ?? 0) < 0 ? "text-rose-600 font-bold" : textSecondary}`}>All: {allBal !== undefined ? Math.round(allBal).toLocaleString() : '—'}</div>
                                        </div>
                                      );
                                    })()}
                                  </div>
                                </td>
                              );
                            })}
                            <td className="px-2 py-2 text-center">
                              <button onClick={() => openFlyout("material", matId)} className={`p-1 rounded ${isDark?'hover:bg-gray-600':'hover:bg-gray-100'}`}><CaretRightIcon /></button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </>
        )}
      </main>

      {/* FLYOUT SIDEBAR */}
      {flyoutOpen && (
        <>
          <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-40" onClick={closeFlyout} />
          <aside className={`fixed right-0 top-0 bottom-0 w-[420px] max-w-full ${bgCard} shadow-2xl z-50 flex flex-col border-l ${borderColor}`}>
            <header className={`px-5 py-4 border-b ${borderColor} flex items-center justify-between ${isDark?'bg-gray-800':'bg-slate-50'}`}>
              <div className="flex-1 pr-4">
                <h2 className="text-lg font-black leading-tight truncate max-w-[280px] text-slate-800 dark:text-slate-100">{flyoutProduct?.name || flyoutMaterial?.name}</h2>
                <span className={`text-[10px] font-bold uppercase tracking-widest ${isDark?'text-gray-400':'text-slate-500'} flex items-center gap-1.5 mt-1`}>
                  {flyoutType === "product" ? <CubeIcon weight="bold" /> : <PackageIcon weight="bold" />} {flyoutType === "product" ? "Product" : "Material"}
                </span>
              </div>
              <button onClick={closeFlyout} className={`p-2 rounded-full hover:bg-gray-200 dark:hover:bg-gray-700 text-gray-400 transition-colors`}><XIcon className="w-5 h-5" /></button>
            </header>

            <div className={`grid grid-cols-2 md:grid-cols-4 gap-3 px-5 py-4 border-b ${borderColor} ${bgCard}`}>
              <div className="flex flex-col">
                <span className={`text-[9px] uppercase font-bold tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-0.5`}>Demand</span>
                <span className="text-base font-black text-orange-600 tabular-nums">{Math.round(flyoutDemand).toLocaleString()}</span>
              </div>
              <div className="flex flex-col">
                <span className={`text-[9px] uppercase font-bold tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-0.5`}>Supply</span>
                <span className="text-base font-black text-blue-600 tabular-nums">{flyoutType === 'material' ? Math.round(flyoutSupply).toLocaleString() : '—'}</span>
              </div>
              <div className="flex flex-col">
                <span className={`text-[9px] uppercase font-bold tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-0.5`}>On Hand</span>
                <span className={`text-base font-black tabular-nums ${textPrimary}`}>{flyoutType === 'material' && flyoutMaterial ? Math.round(flyoutMaterial.onHand).toLocaleString() : '—'}</span>
              </div>
              <div className="flex flex-col">
                <span className="text-[9px] uppercase font-bold tracking-widest text-emerald-600 dark:text-emerald-50 mb-0.5">Avail Now</span>
                {flyoutType === 'material' && flyoutMaterial ? (() => {
                    const availNow = availableNowSandboxByMaterial.get(flyoutId) ?? 0;
                    return (
                        <span className={`text-base font-black tabular-nums ${availNow >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400'}`}>
                          {Math.round(availNow).toLocaleString()}
                        </span>
                    );
                })() : <span className={`text-base font-black ${textPrimary}`}>—</span>}
              </div>
            </div>

            {flyoutType === "material" && flyoutMaterial && (() => {
              const committedQty = totalCommittedSandboxByMaterial.get(flyoutId) ?? 0;
              let fReady = 0;
              let fOnOrder = 0;
              if (flyoutMaterial.orderedStr) {
                const regex = /([\d,.]+)\s*(ready|on order)/gi;
                let match;
                while ((match = regex.exec(flyoutMaterial.orderedStr)) !== null) {
                  const numStr = match[1].replace(/[^\d.]/g, '');
                  const qty = parseFloat(numStr) || 0;
                  if (match[2].toLowerCase() === 'ready') fReady += qty;
                  else if (match[2].toLowerCase() === 'on order') fOnOrder += qty;
                }
              }
              return (
                <div className={`px-5 py-3 border-b ${borderColor} ${clsFlyoutSubBg} shrink-0`}>
                  <div className="grid grid-cols-3 gap-3">
                     <div className="flex flex-col">
                        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-transparent border border-indigo-300 text-indigo-700 dark:border-indigo-700 dark:text-indigo-400 w-max mb-1">On Order</span>
                        <span className="text-sm font-black dark:text-gray-100 tabular-nums text-slate-700 dark:text-slate-200">{fOnOrder > 0 ? Math.round(fOnOrder).toLocaleString() : '—'}</span>
                     </div>
                     <div className="flex flex-col">
                        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-blue-50 border border-blue-200 text-blue-700 dark:bg-blue-900/30 dark:border-blue-800 dark:text-blue-400 w-max mb-1">Ready</span>
                        <span className="text-sm font-black dark:text-gray-100 tabular-nums text-slate-700 dark:text-slate-200">{fReady > 0 ? Math.round(fReady).toLocaleString() : '—'}</span>
                     </div>
                     <div className="flex flex-col">
                        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider bg-transparent border border-orange-300 text-orange-700 dark:border-orange-700 dark:text-orange-400 w-max mb-1">Committed</span>
                        <span className="text-sm font-black text-slate-700 dark:text-slate-200 tabular-nums">{committedQty > 0 ? `-${Math.round(committedQty).toLocaleString()}` : '—'}</span>
                     </div>
                  </div>
                </div>
              );
            })()}

            {flyoutType === "material" && (
              <div className={`flex flex-col border-b ${borderColor} ${bgCard} shrink-0`}>
                <div className={`px-5 py-3 flex justify-between items-center ${clsFlyoutRowBg}`}>
                  <div className="flex flex-col">
                    <span className="text-xs font-bold text-slate-600 dark:text-slate-300">Need more stock?</span>
                    {!(suggestedTable && suggestedMaterialField?.id && suggestedQtyField?.id) && (
                      <span className="text-[9px] text-amber-600 mt-0.5 uppercase tracking-widest font-bold">Map 'Suggested Buys' in settings</span>
                    )}
                  </div>
                  <button
                    onClick={() => {
                      if (suggestedTable && suggestedMaterialField?.id && suggestedQtyField?.id) {
                        setSuggestBuyModal({ materialId: flyoutId, qty: 100 });
                      } else {
                        addToast('error', 'Configure the Suggested Buys table first.');
                      }
                    }}
                    className={`flex items-center gap-1 px-3 py-1.5 text-[10px] font-black uppercase tracking-widest rounded border border-indigo-200 bg-indigo-50 text-indigo-600 dark:bg-indigo-900/50 dark:border-indigo-800 dark:text-indigo-400 transition-colors shadow-sm hover:bg-indigo-100 dark:hover:bg-indigo-900`}
                  >
                    <PlusIcon weight="bold" /> Request
                  </button>
                </div>
                {suggestedBuys.filter(b => b.materialId === flyoutId && (!selectedPlanId || b.planId === selectedPlanId)).length > 0 && (
                  <div className={`px-5 pb-3 pt-1 flex flex-col gap-2 ${clsFlyoutRowBg}`}>
                    {suggestedBuys.filter(b => b.materialId === flyoutId && (!selectedPlanId || b.planId === selectedPlanId)).map(buy => (
                      <div key={buy.id} className={`flex justify-between items-center p-3 rounded-xl border ${clsSuggestCard} shadow-sm`}>
                        <div className="flex items-center gap-3">
                          <span className={`text-sm font-black tabular-nums ${isDark?'text-indigo-400':'text-indigo-700'}`}>{Math.round(buy.qty).toLocaleString()} UNITS</span>
                          <span className={`text-[9px] uppercase font-bold ${isDark ? 'text-indigo-500' : 'text-indigo-400'} tracking-widest px-2 py-0.5 rounded-full bg-indigo-100`}>Pending Review</span>
                        </div>
                        <button onClick={() => setSuggestBuyModal({ id: buy.id, materialId: buy.materialId, qty: buy.qty })} className={`p-1.5 ${isDark?'text-indigo-400 hover:bg-indigo-800':'text-indigo-500 hover:bg-indigo-200'} rounded-lg transition-colors`}>
                          <PencilSimpleIcon size={16} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            <div className={`flex-1 overflow-y-auto p-5 space-y-8 ${bgCard}`}>
              {flyoutType === "material" && (
                <>
                  <div>
                    <h3 className={`text-xs font-black uppercase tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-3 flex items-center gap-2`}><Truck size={16} className="text-blue-500" /> Purchase Orders</h3>
                    <div className="space-y-3">
                      {poGroups.length === 0 ? <div className={`text-sm ${isDark?'text-gray-500':'text-slate-400'} italic px-2`}>No linked POs.</div> : poGroups.map((group) => (
                        <div key={group.poLineId} className={`border ${borderColor} rounded-xl shadow-sm ${isDark?'bg-gray-800':'bg-white'} overflow-hidden`}>
                          <div className={`p-4 flex justify-between items-start ${isDark?'bg-gray-800':'bg-slate-50'}`}>
                            <div className="flex flex-col">
                              <span className="text-sm font-black text-slate-900 dark:text-slate-100">PO {group.name}</span>
                              <span className={`text-[10px] ${textSecondary} uppercase flex items-center gap-1.5 mt-1`}>
                                <span className={`px-1.5 py-0.5 rounded border ${clsPoStatusBadge} text-[8px] font-bold tracking-widest leading-none`}>{group.status}</span>
                                <span className="font-bold">{group.vendor}</span>
                              </span>
                            </div>
                            <div className="flex flex-col items-end text-right">
                              <span className={`font-mono font-black ${isDark?'text-blue-400':'text-blue-700'} text-sm`}>{group.remaining ? Math.round(group.remaining.qty).toLocaleString() : '0'}</span>
                              <span className={`${textSecondary} font-medium flex items-center gap-1 text-[10px] mt-0.5 uppercase tracking-widest font-bold`}>UNSCHEDULED</span>
                            </div>
                          </div>
                          {group.shipments.length > 0 && (
                            <div className={`border-t ${borderColor} ${clsShipmentsBg} flex flex-col`}>
                              {group.shipments.map((s: any) => (
                                <div key={s.id} className={`px-4 py-2.5 flex justify-between items-center border-b ${borderColor} last:border-0 hover:bg-black/5 dark:hover:bg-white/5 transition-colors`}>
                                  <div className="flex items-center gap-2">
                                    <CaretRightIcon className={`${textSecondary} w-3.5 h-3.5`} weight="bold" />
                                    <span className={`text-[10px] uppercase font-bold tracking-wider ${isDark?'text-slate-300':'text-slate-600'}`}>{s.shipmentName || 'Shipment'}</span>
                                  </div>
                                  <div className="flex items-center gap-5">
                                    <span className={`text-xs ${textSecondary} font-medium`}>{formatDate(s.date)}</span>
                                    <span className={`font-mono font-black text-sm ${isDark?'text-indigo-400':'text-indigo-600'} w-14 text-right`}>{Math.round(s.qty).toLocaleString()}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                          <div className={`p-2 border-t ${borderColor} ${isDark?'bg-gray-800':'bg-slate-50'}`}>
                            <button onClick={(e) => expandPO(e, group.headerId)} className="w-full block text-center text-[10px] uppercase font-bold tracking-widest text-blue-600 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-300 bg-blue-50/50 hover:bg-blue-100 dark:bg-blue-900/20 dark:hover:bg-blue-900/50 py-2 rounded-lg border border-blue-100 dark:border-blue-800 transition-colors">
                              Open Purchase Order ↗
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div>
                    <h3 className={`text-xs font-black uppercase tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-3 flex items-center gap-2`}><CircleIcon weight="fill" className="text-purple-500" /> Virtual POs</h3>
                    <div className="space-y-3">
                      {flyoutVPOs.length === 0 ? <div className={`text-sm ${isDark?'text-gray-500':'text-slate-400'} italic px-2`}>No virtual POs.</div> : flyoutVPOs.map((vpo) => (
                        <div key={vpo.id} className={`p-4 border ${clsVpoBorder} rounded-xl shadow-sm`}>
                          <div className="flex justify-between items-start mb-3">
                            <div className="flex flex-col">
                              <span className={`text-sm font-black ${isDark?'text-purple-400':'text-purple-700'}`}>Simulated PO</span>
                              <span className={`text-[10px] ${isDark?'text-gray-400':'text-slate-500'} uppercase font-bold tracking-wider mt-0.5`}>{vpo.vendor}</span>
                            </div>
                            <div className="flex gap-1.5">
                              <button onClick={() => { setEditingVpo(vpo); setVpoForm({ materialId: vpo.materialId, qty: vpo.qty, date: vpo.date, vendor: vpo.vendor }); }} className={`p-1.5 ${clsVpoEditBtn} rounded-lg transition-colors`}><PencilSimpleIcon size={16}/></button>
                              <button onClick={() => deleteVpo(vpo.id)} className={`p-1.5 ${clsVpoDeleteBtn} rounded-lg transition-colors`}><TrashIcon size={16}/></button>
                            </div>
                          </div>
                          <div className={`flex justify-between items-center text-xs pt-3 border-t ${clsVpoDivider}`}>
                            <span className={`font-mono font-black ${clsVpoQtyBadge} px-2 py-0.5 rounded text-sm`}>+{Math.round(vpo.qty).toLocaleString()} UNITS</span>
                            <span className={`${isDark?'text-gray-400':'text-slate-500'} font-medium flex items-center gap-1.5 text-[11px]`}><Calendar size={14} /> {formatDate(vpo.date)}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </>
              )}

              <div>
                <h3 className={`text-xs font-black uppercase tracking-widest ${isDark?'text-gray-400':'text-slate-400'} mb-3 flex items-center gap-2`}><ShoppingCart size={16} className="text-orange-500" /> Linked Sales Orders</h3>
                <div className="space-y-3">
                  {flyoutSOs.length === 0 ? <div className={`text-sm ${isDark?'text-gray-500':'text-slate-400'} italic px-2`}>No linked orders.</div> : flyoutSOs.map((so) => {
                    const lis = getLineItemsForOrder(so.id);
                    return (
                      <div key={so.id} className={`p-4 border ${borderColor} rounded-xl shadow-sm ${isDark?'bg-gray-800':'bg-white'}`}>
                        <div className="flex justify-between items-start mb-4">
                          <div className="flex flex-col pr-3">
                            <span className="text-sm font-black text-slate-900 dark:text-slate-100">{so.name}</span>
                            <span className={`text-[10px] ${isDark?'text-gray-400':'text-slate-500'} uppercase font-bold tracking-wider mt-0.5`}>{so.customer} {so.customerPO ? `• PO: ${so.customerPO}` : ''} • {formatDate(so.date)}</span>
                          </div>
                          <select value={so.soStatus || 'Pending'} onChange={(e) => updateSOStatus(so.id, e.target.value)} className={`shrink-0 px-2 py-1 rounded text-[9px] font-black uppercase tracking-widest border ${borderColor} ${isDark?'bg-gray-700 text-gray-200':'bg-slate-50 text-slate-700'} cursor-pointer hover:border-indigo-400 focus:ring-2 focus:ring-indigo-500 outline-none transition-colors`}>
                            {soStatusOptions.map(o => <option key={o} value={o}>{o}</option>)}
                          </select>
                        </div>
                        <div className="flex flex-col gap-3 mb-2">
                          {lis.map((li, idx) => {
                             const product = products.find(p => p.id === li.productId);
                             const prodName = product ? product.name : "Unknown Product";
                             if (flyoutType === "material") {
                               const bomsForMaterial = bom.filter(b => b.productId === li.productId && b.materialId === flyoutId);
                               const usesMaterial = bomsForMaterial.length > 0;
                               if (usesMaterial) {
                                 return (
                                   <div key={idx} className="flex flex-col">
                                     <div className="flex justify-between items-center">
                                       <span className="text-xs font-black text-slate-900 dark:text-slate-100">{prodName}</span>
                                       <span className="text-xs font-black tabular-nums">{Math.round(li.qty).toLocaleString()} Units</span>
                                     </div>
                                     {bomsForMaterial.map((b, i) => {
                                        const reqQty = b.qtyPer * li.qty;
                                        const wf = allSandboxWaterfalls.get(flyoutId); 
                                        const dData = wf?.get(so.date);
                                        const bal = dData?.balance ?? (availableNowSandboxByMaterial.get(flyoutId) ?? 0);
                                        const short = bal < 0 ? Math.min(reqQty, Math.abs(bal)) : 0;
                                        return (
                                          <div key={i} className="pl-3 mt-1.5 mb-1.5 flex justify-between items-center text-[10px] border-l-2 border-indigo-200 dark:border-indigo-800">
                                            <span className={`${textSecondary} font-bold`}>↳ Requires: {Math.round(reqQty).toLocaleString()} {flyoutMaterial?.name}</span>
                                            <span className={`font-black ${short > 0 ? 'text-rose-600' : 'text-emerald-600'}`}>{short > 0 ? `(Short: ${Math.round(short).toLocaleString()})` : '(OK)'}</span>
                                          </div>
                                        );
                                     })}
                                   </div>
                                 );
                               } else {
                                 return (
                                   <div key={idx} className={`flex justify-between items-center opacity-40 hover:opacity-70 transition-opacity`}>
                                     <span className="text-xs font-medium text-slate-500 dark:text-slate-400">{prodName}</span>
                                     <span className="text-xs font-medium tabular-nums text-slate-500 dark:text-slate-400">{Math.round(li.qty).toLocaleString()} Units</span>
                                   </div>
                                 );
                               }
                             } else {
                               const isClickedProduct = li.productId === flyoutId;
                               return (
                                 <div key={idx} className={`flex justify-between items-center ${isClickedProduct ? 'font-black text-slate-900 dark:text-slate-100' : 'opacity-40 hover:opacity-70 transition-opacity text-slate-500 dark:text-slate-400'}`}>
                                   <span className="text-xs">{prodName}</span>
                                   <span className="text-xs tabular-nums">{Math.round(li.qty).toLocaleString()} Units</span>
                                 </div>
                               );
                             }
                          })}
                        </div>
                        <button onClick={(e) => expandSO(e, so.id)} className="w-full mt-3 block text-center text-[10px] uppercase font-bold tracking-widest text-indigo-600 hover:text-indigo-700 dark:text-indigo-400 dark:hover:text-indigo-300 bg-indigo-50 hover:bg-indigo-100 dark:bg-indigo-900/30 dark:hover:bg-indigo-900/50 py-2 rounded-lg border border-indigo-100 dark:border-indigo-800 transition-colors">
                          Open Sales Order ↗
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </aside>
        </>
      )}

      {/* SUGGEST BUY MODAL */}
      {suggestBuyModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
          <div className={`w-full max-w-md rounded-2xl shadow-2xl ${bgCard} p-6 border ${borderColor}`}>
            <div className={`flex items-center justify-between mb-5 border-b ${borderColor} pb-4`}>
              <h2 className="text-lg font-black uppercase tracking-tight flex items-center gap-2"><CheckCircle size={20} className="text-indigo-500" /> {suggestBuyModal.id ? 'Update Request' : 'Suggest Buy'}</h2>
              <button onClick={() => setSuggestBuyModal(null)} className={`p-1.5 rounded-md ${hoverBgButton}`}><XIcon className="w-5 h-5" /></button>
            </div>
            <div className="space-y-4">
              <div>
                <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Material</label>
                <div className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${isDark?'bg-gray-800':'bg-slate-50'} font-bold`}>{materials.find(m => m.id === suggestBuyModal.materialId)?.name}</div>
              </div>
              {selectedPlan && (
                <div>
                  <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Linked Plan</label>
                  <div className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${isDark?'bg-gray-800':'bg-slate-50'} font-bold`}>{selectedPlan.name}</div>
                </div>
              )}
              <div>
                <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Quantity Needed</label>
                <input type="number" value={suggestBuyModal.qty} onChange={(e) => setSuggestBuyModal(p => p ? { ...p, qty: parseInt(e.target.value) || 0 } : null)} className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${bgCard} font-bold focus:ring-2 focus:ring-indigo-500 outline-none`} />
              </div>
              <div className="flex gap-4 pt-4 mt-2">
                <button onClick={() => setSuggestBuyModal(null)} className={`flex-1 py-3 rounded-xl text-xs font-black uppercase tracking-widest ${isDark?'bg-gray-700 hover:bg-gray-600':'bg-slate-100 hover:bg-slate-200'}`}>Cancel</button>
                <button onClick={submitBuySuggestion} disabled={suggestBuyModal.qty <= 0} className="flex-1 py-3 rounded-xl text-xs font-black uppercase tracking-widest bg-indigo-600 text-white hover:bg-indigo-700 disabled:opacity-50">Send Request</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* VPO MODAL */}
      {editingVpo && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
          <div className={`w-full max-w-md rounded-2xl shadow-2xl ${bgCard} p-6 border ${borderColor}`}>
            <div className={`flex items-center justify-between mb-5 border-b ${borderColor} pb-4`}>
              <h2 className="text-lg font-black uppercase tracking-tight flex items-center gap-2"><CubeIcon size={20} className="text-purple-500" /> {editingVpo.id ? "Edit Virtual PO" : "Simulate Purchase"}</h2>
              <button onClick={() => setEditingVpo(null)} className={`p-1.5 rounded-md ${hoverBgButton}`}><XIcon className="w-5 h-5" /></button>
            </div>
            <div className="space-y-4">
              <div>
                <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Material</label>
                <select value={vpoForm.materialId} onChange={(e) => setVpoForm((p) => ({ ...p, materialId: e.target.value }))} className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${bgCard} font-bold focus:ring-2 focus:ring-purple-500 outline-none`}>
                  {materials.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Quantity</label>
                  <input type="number" value={vpoForm.qty} onChange={(e) => setVpoForm((p) => ({ ...p, qty: parseInt(e.target.value) || 0 }))} className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${bgCard} font-bold focus:ring-2 focus:ring-purple-500 outline-none`} />
                </div>
                <div>
                  <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>ETA Date</label>
                  <input type="date" value={vpoForm.date} onChange={(e) => setVpoForm((p) => ({ ...p, date: e.target.value }))} className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${bgCard} font-bold focus:ring-2 focus:ring-purple-500 outline-none`} />
                </div>
              </div>
              <div>
                <label className={`block text-xs uppercase font-bold tracking-widest ${textSecondary} mb-1.5`}>Vendor Note</label>
                <input type="text" value={vpoForm.vendor} onChange={(e) => setVpoForm((p) => ({ ...p, vendor: e.target.value }))} className={`w-full px-4 py-3 rounded-xl border ${borderColor} ${bgCard} font-bold focus:ring-2 focus:ring-purple-500 outline-none`} />
              </div>
              <div className="flex gap-4 pt-4 mt-2">
                <button onClick={() => setEditingVpo(null)} className={`flex-1 py-3 rounded-xl text-xs font-black uppercase tracking-widest ${isDark?'bg-gray-700 hover:bg-gray-600':'bg-slate-100 hover:bg-slate-200'}`}>Cancel</button>
                <button onClick={saveVpo} className="flex-1 py-3 rounded-xl text-xs font-black uppercase tracking-widest bg-purple-600 text-white hover:bg-purple-700">Save Simulation</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* CLICK INFO POPUPS */}
      {popupInfo && (
        <>
          <div className="fixed inset-0 z-[60]" onClick={closePopup} />
          <div
            className={`fixed z-[70] rounded-2xl shadow-2xl ${bgCard} border ${borderColor} flex flex-col overflow-hidden`}
            style={{
              left: popupInfo.x,
              width: 420,
              ...(popupInfo.top !== undefined ? { top: popupInfo.top } : {}),
              ...(popupInfo.bottom !== undefined ? { bottom: popupInfo.bottom } : {}),
              maxHeight: popupInfo.maxHeight || 520,
            }}
            role="dialog"
          >
            <div className="flex-1 overflow-y-auto p-4">
              {popupInfo.content}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

initializeBlock({ interface: () => <SupplyChainPlanner /> });
