import { Reshuffle, BaseHttpConnector } from 'reshuffle-base-connector'
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client'

import { ClientCredentialAuthenticationProvider, MicrosoftOptions } from './auth'

export type MicrosoftConnectorConfigOptions = MicrosoftOptions & { debugLogging?: boolean }
export class MicrosoftConnector extends BaseHttpConnector<
  MicrosoftConnectorConfigOptions,
  undefined
> {
  client: Client

  constructor(app: Reshuffle, options: MicrosoftConnectorConfigOptions, id?: string) {
    super(app, options, id)
    const { debugLogging = false, ...authOptions } = options
    const clientOptions: ClientOptions = {
      defaultVersion: 'v1.0',
      debugLogging,
      authProvider: new ClientCredentialAuthenticationProvider(authOptions),
    }
    this.client = Client.initWithMiddleware(clientOptions)
  }

  sdk(): Client {
    return this.client
  }
}
