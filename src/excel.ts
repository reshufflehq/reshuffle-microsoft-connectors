import './fetch'
import MicrosoftGraph from '@microsoft/microsoft-graph-types'

import { MicrosoftConnector } from './base'

export default class ExcelConnector extends MicrosoftConnector {
  async getDriveItem(driveItem: string): Promise<MicrosoftGraph.DriveItem> {
    const request = await this.client.api(driveItem).get()
    return request
  }

  async listWorksheets(driveItem: string): Promise<MicrosoftGraph.WorkbookWorksheet[]> {
    const request = await this.client.api(`${driveItem}/workbook/worksheets`).get()
    return request.value
  }

  async getTables(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookTable[]> {
    const request = await this.client.api(`${driveItem}/workbook/worksheets/${name}/tables`).get()
    return request.value
  }

  async getCharts(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookChart[]> {
    const request = await this.client.api(`${driveItem}/workbook/worksheets/${name}/charts`).get()
    return request.value
  }

  async addNewWorksheet(
    driveItem: string,
    name: string,
  ): Promise<MicrosoftGraph.WorkbookWorksheet> {
    const request = await this.client.api(`${driveItem}/workbook/worksheets`).post({ name })
    return request
  }

  async getWorksheet(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookWorksheet> {
    const request = await this.client.api(`${driveItem}/workbook/worksheets/${name}`).get()
    return request
  }

  async updateWorksheet(
    driveItem: string,
    name: string,
    newName?: string,
    position?: number,
    visibility?: 'Visible' | 'Hidden' | 'VeryHidden',
  ): Promise<MicrosoftGraph.WorkbookWorksheet> {
    const request = await this.client
      .api(`${driveItem}/workbook/worksheets/${name}`)
      .update({ position, name: newName, visibility })
    return request
  }

  async getCell(
    driveItem: string,
    name: string,
    row: number,
    column: number,
  ): Promise<MicrosoftGraph.WorkbookRange> {
    const request = await this.client
      .api(`${driveItem}/workbook/worksheets/${name}/cell(row=${row},column=${column})`)
      .get()
    return request
  }

  async getRange(
    driveItem: string,
    name: string,
    address: string,
  ): Promise<MicrosoftGraph.WorkbookRange> {
    const request = await this.client
      .api(`${driveItem}/workbook/worksheets/${name}/range(address='${address}')`)
      .get()
    return request
  }

  async updateRange(
    driveItem: string,
    name: string,
    address: string,
    values?: Record<string, unknown>,
    formula?: Record<string, unknown>,
    numberFormat?: Record<string, unknown>,
  ): Promise<MicrosoftGraph.WorkbookRange> {
    const request = await this.client
      .api(`${driveItem}/workbook/worksheets/${name}/range(address='${address}')`)
      .update({ values, formula, numberFormat })
    return request
  }

  async insertRange(
    driveItem: string,
    name: string,
    values?: Record<string, unknown>,
    formula?: Record<string, unknown>,
    numberFormat?: Record<string, unknown>,
  ): Promise<MicrosoftGraph.WorkbookRange> {
    const request = await this.client
      .api(`${driveItem}/workbook/worksheets/${name}/range/insert`)
      .update({ values, formula, numberFormat })
    return request
  }
}

export { ExcelConnector }
