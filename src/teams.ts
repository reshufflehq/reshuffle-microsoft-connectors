import './fetch'
import type MicrosoftGraph from '@microsoft/microsoft-graph-types'

import { MicrosoftConnector } from './base'

export class TeamsConnector extends MicrosoftConnector {
  // To get all teams you use groups
  async listGroups(): Promise<MicrosoftGraph.Group[]> {
    const request = await this.client
      .api('/groups')
      .version('beta')
      .filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
      .get()
    return request.value
  }

  async listChannels(teamId: string): Promise<MicrosoftGraph.Channel[]> {
    const request = await this.client.api(`/teams/${teamId}/channels`).version('beta').get()
    return request.value
  }
}
