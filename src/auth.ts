import './fetch'
import { AuthenticationProvider } from '@microsoft/microsoft-graph-client'
import { URLSearchParams } from 'url'

export interface MicrosoftOptions {
  AppId: string
  AppPassword: string
  AppTenantId: string
}

export class ClientCredentialAuthenticationProvider implements AuthenticationProvider {
  options: MicrosoftOptions

  constructor(options: MicrosoftOptions) {
    this.options = options
  }

  public async getAccessToken(): Promise<string> {
    const response = await fetch(
      `https://login.microsoftonline.com/${this.options.AppTenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        body: new URLSearchParams({
          client_id: this.options.AppId,
          client_secret: this.options.AppPassword,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      },
    )
    const data = await response.json()
    if (response.ok) {
      return data.access_token
    } else {
      throw new Error(`${data.error} - ${data.error_description}`)
    }
  }
}
