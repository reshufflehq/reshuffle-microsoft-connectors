# reshuffle-teams-connector

`npm install reshuffle-microsoft-connectors`

_ES6 import_: `import { TeamsConnector } from 'reshuffle-microsoft-connectors'`

This is a [Reshuffle](https://reshuffle.com) connector that provides an Interface to Microsoft Teams.

The following example adds a new worksheet to an excel file

```js
const { Reshuffle } = require('reshuffle')
const { TeamsConnector } = require('reshuffle-microsoft-connectors')

const app = new Reshuffle()
const connector = new TeamsConnector(app, { process.env.AppId, process.env.AppPassword, process.env.AppTenantId })

connector.on(
  {
    resource: '/teams/getAllMessages',
    changeType: 'created',
    runtimeBaseUrl: 'https://example.com',
    expirationDateTime: '2021-02-03T03:47:17.292Z',
  },
  (event) => console.log(event),
)

connector.listTeams().then((teams) => console.log(teams))
```

#### Table of Contents

[Configuration options](#Configuration-Options)

[TypeScript Types](#TypeScript-Types)

_Connector actions_:

[listTeams](#listTeams) List Teams

[listChannels](#listChannels) List Channels

##### Configuration options

```js
const app = new Reshuffle()
const connector = new TeamsConnector(app, { AppId, AppPassword, AppTenantId })
```

Credentials can be created by following the guide at https://docs.microsoft.com/en-us/graph/auth-v2-service

See the `Credentials` interface exported from the connector for details.

##### TypeScript types

The following types are exported from the connector:

- **interface MicrosoftOptions** Microsoft Credentials

#### Connector actions

##### listGroups

List [Groups](https://docs.microsoft.com/en-us/graph/api/resources/group).

```ts
async listGroups(): Promise<MicrosoftGraph.Group[]>
```

##### listChannels

List [Channels](https://docs.microsoft.com/en-us/graph/api/resources/channel).

```ts
async listChannels(): Promise<MicrosoftGraph.Channel[]>
```
