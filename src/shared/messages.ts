export type Message =
  | { type: 'SCAN_PAGE' }
  | { type: 'PREVIEW'; requirements: import('./types').ExcelRequirement[] }
  | { type: 'FILL'; requirements: import('./types').ExcelRequirement[]; options: import('./types').FillOptions }
  | { type: 'DEBUG_DOM' }
  | { type: 'INSTALL_HELPERS'; requirements: import('./types').ExcelRequirement[] }
  | { type: 'SET_DESCRIPTION_COLOR'; enabled: boolean }
  | { type: 'STORE_WORKBOOK'; procedureId: string; fileName: string; requirements: import('./types').ExcelRequirement[] }
  | { type: 'GET_WORKBOOK'; procedureId: string }
  | { type: 'CLEAR_WORKBOOK'; procedureId: string };
