# reshuffle-excel-connector

[Code](https://github.com/reshufflehq/reshuffle-microsoft-connectors) |
[npm](https://www.npmjs.com/package/reshuffle-microsoft-connectors) |
[Code sample](https://github.com/reshufflehq/reshuffle/tree/master/examples/microsoft/excel)

`npm install reshuffle-excel-connector`

### Reshuffle Microsoft Connectors

This package contains a collection of [Reshuffle](https://github.com/reshufflehq/reshuffle)
connectors to [Microsoft](https://microsoft.com).

The following example adds a new worksheet to an excel file

```js
const { Reshuffle } = require('reshuffle')
const { MicrosoftConnector } = require('reshuffle-microsoft-connectors')

const app = new Reshuffle()
const connector = new ExcelConnector(app, { AppId, AppPassword, AppTenantId })

excelConnector
  .addNewWorksheet('drive/items/{item-id}', 'myNewSheet')
  .then((newWorksheet) => console.log(newWorksheet))
```

#### Table of Contents

[Configuration](#ConfigurationOptions) Configuration options

[Types](#TypeScriptTypes) TypeScript types

_Connector actions_:

[getDriveItem](#getDriveItem) Get a DriveItem

[listWorksheets](#listWorksheets) List Worksheets

[getTables](#getTables) Get Tables

[getCharts](#getCharts) Get Charts

[addNewWorksheet](#addNewWorksheet) Add a Worksheet

[getWorksheet](#addWorksheet) Get a Worksheet

[updateWorksheet](#updateWorksheet) Update a Worksheet

[getCell](#getCell) Get a cell

[getRange](#getRange) Get a range

[updateRange](#updateRange) Update a range

[insertRange](#insertRange) Insert a range

##### Configuration options

```js
const app = new Reshuffle()
const connector = new ExcelConnector(app, { AppId, AppPassword, AppTenantId })
```

Credentials can be created by following the guide at https://docs.microsoft.com/en-us/graph/auth-v2-service

See the `Credentials` interface exported from the connector for details.

##### TypeScript types

The following types are exported from the connector:

- **interface MicrosoftOptions** Microsoft Credentials

#### Connector actions

##### getDriveItem

Get a [DriveItem](https://docs.microsoft.com/en-us/graph/api/resources/driveitem).

```ts
async getDriveItem(driveItem: string): Promise<MicrosoftGraph.DriveItem>
```

##### listWorksheets

List Worksheets.

```ts
async listWorksheets(driveItem: string): Promise<MicrosoftGraph.WorkbookWorksheet[]>
```

##### getTables

Get Tables.

```ts
async getTables(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookTable[]>
```

##### getTables

Get Tables.

```ts
async getTables(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookTable[]>
```

##### getCharts

Get Charts.

```ts
async getCharts(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookChart[]>
```

##### addNewWorksheet

Add a new Worksheet.

```ts
async addNewWorksheet(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookWorksheet>
```

##### getWorksheet

Get a Worksheet.

```ts
async getWorksheet(driveItem: string, name: string): Promise<MicrosoftGraph.WorkbookWorksheet>
```

##### updateWorksheet

Update a Worksheet.

```ts
async updateWorksheet(
    driveItem: string,
    name: string,
    newName?: string,
    position?: number,
    visibility?: 'Visible' | 'Hidden' | 'VeryHidden',
  ): Promise<MicrosoftGraph.WorkbookWorksheet>
```

##### getCell

Get a cell.

```ts
async getCell(
    driveItem: string,
    name: string,
    row: number,
    column: number,
  ): Promise<MicrosoftGraph.WorkbookRange>
```

##### getRange

Get a Range.

```ts
async getRange(
    driveItem: string,
    name: string,
    address: string,
  ): Promise<MicrosoftGraph.WorkbookRange>
```

##### updateRange

Update a Range.

```ts
async updateRange(
    driveItem: string,
    name: string,
    address: string,
    values?: Record<string, unknown>,
    formula?: Record<string, unknown>,
    numberFormat?: Record<string, unknown>,
  ): Promise<MicrosoftGraph.WorkbookRange>
```

##### insertRange

Insert a Range.

```ts
async insertRange(
    driveItem: string,
    name: string,
    values?: Record<string, unknown>,
    formula?: Record<string, unknown>,
    numberFormat?: Record<string, unknown>,
  ): Promise<MicrosoftGraph.WorkbookRange>
```

##### insertRange

Insert a Range.

```ts
async insertRange(
    driveItem: string,
    name: string,
    values?: Record<string, unknown>,
    formula?: Record<string, unknown>,
    numberFormat?: Record<string, unknown>,
  ): Promise<MicrosoftGraph.WorkbookRange>
```
