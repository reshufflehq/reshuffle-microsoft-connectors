import { Reshuffle, BaseHttpConnector } from 'reshuffle-base-connector'
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client'

import { ClientCredentialAuthenticationProvider, MicrosoftOptions } from './auth'

export type MicrosoftConnectorConfigOptions = MicrosoftOptions

export default class MicrosoftConnector extends BaseHttpConnector<
  MicrosoftConnectorConfigOptions,
  undefined
> {
  client: Client

  constructor(app: Reshuffle, options: MicrosoftConnectorConfigOptions, id?: string) {
    super(app, options, id)
    const clientOptions: ClientOptions = {
      defaultVersion: 'v1.0',
      debugLogging: false,
      authProvider: new ClientCredentialAuthenticationProvider(options),
    }
    this.client = Client.initWithMiddleware(clientOptions)
  }

  sdk(): Client {
    return this.client
  }
}

export { MicrosoftConnector }
