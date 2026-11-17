export interface HwpCell {
  row: number;
  col: number;
  row_span: number;
  col_span: number;
  text: string;
}

export interface HwpTable {
  index: number;
  total_rows: number;
  total_cols: number;
  cells: HwpCell[];
}

export interface ExcelCell {
  row: number;
  col: number;
  merge_rows: number;
  merge_cols: number;
  value: string;
  address: string;
}

export interface ExcelSheet {
  name: string;
  cells: ExcelCell[];
  max_row: number;
  max_col: number;
}

export interface ExcelTemplate {
  sheets: ExcelSheet[];
  path: string;
}

export interface CellMapping {
  hwp_table_idx: number;
  hwp_row: number;
  hwp_col: number;
  excel_sheet: number;
  excel_row: number;
  excel_col: number;
}

export interface SelectedHwpCell {
  tableIdx: number;
  row: number;
  col: number;
}
