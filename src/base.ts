import { Reshuffle, BaseHttpConnector, EventConfiguration } from 'reshuffle-base-connector'
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client'
import type MicrosoftGraph from '@microsoft/microsoft-graph-types'
import type { Request, Response } from 'express'

import { ClientCredentialAuthenticationProvider, MicrosoftOptions } from './auth'

export type MicrosoftConnectorConfigOptions = Omit<MicrosoftOptions, 'notificationUrl'> & {
  debugLogging?: boolean
  runtimeBaseUrl?: string
  webhookPath?: string
}

export type MicrosoftEventConfigOptions = MicrosoftGraph.Subscription

const DEFAULT_WEBHOOK_PATH = '/webhooks/microsoft'

export class MicrosoftConnector extends BaseHttpConnector<
  MicrosoftConnectorConfigOptions,
  MicrosoftEventConfigOptions
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

  async onStart(): Promise<void> {
    const logger = this.app.getLogger()
    const subscriptions: MicrosoftGraph.Subscription[] = (
      await this.sdk().api('/subscriptions').get()
    ).value

    for (const event of Object.values(this.eventConfigurations)) {
      const { runtimeBaseUrl, ...options } = event.options as MicrosoftConnectorConfigOptions
      const eventSubscription = {
        ...options,
        notificationUrl: runtimeBaseUrl + (DEFAULT_WEBHOOK_PATH || options.webhookPath),
      } as MicrosoftGraph.Subscription
      const existingSubscription = subscriptions.find((subscription) =>
        Object.keys(event.options).every(
          (key) =>
            eventSubscription[key as keyof MicrosoftGraph.Subscription] ===
            subscription[key as keyof MicrosoftGraph.Subscription],
        ),
      )
      let subscription: MicrosoftGraph.Subscription
      if (existingSubscription) {
        logger.info(
          `Reshuffle Microsoft - existing webhook reused (resource: ${existingSubscription.resource}, url: ${existingSubscription.notificationUrl})`,
        )
        subscription = existingSubscription
      } else {
        subscription = await this.sdk().api('/subscriptions').post(eventSubscription)
        logger.info(
          `Reshuffle Microsoft - webhook registered successfully (resource: ${subscription.resource}, url: ${subscription.notificationUrl})`,
        )
      }
      this.eventConfigurations[event.id].options = {
        ...this.eventConfigurations[event.id].options,
        subscriptionId: subscription.id,
      }
    }
  }

  // Your events
  on(
    options: MicrosoftEventConfigOptions,
    handler: (event: EventConfiguration & Record<string, any>, app: Reshuffle) => void,
    eventId: string,
  ): EventConfiguration {
    const path = this.configOptions?.webhookPath || DEFAULT_WEBHOOK_PATH

    if (!eventId) {
      eventId = `Microsoft${path}/${options.resource}/${this.id}`
    }
    const event = new EventConfiguration(eventId, this, options)
    this.eventConfigurations[event.id] = event

    this.app.when(event, handler as any)
    this.app.registerHTTPDelegate(path, this)

    return event
  }

  async handle(req: Request, res: Response): Promise<boolean> {
    const { validationToken } = req.query

    if (validationToken) {
      res.send(validationToken)
    } else {
      res.status(202).send()
      const { value: data = [] } = req.body as MicrosoftGraph.ChangeNotificationCollection
      for (const incomingEvent of data) {
        const eventsUsingMicrosoftEvent = Object.values(this.eventConfigurations).filter(
          (event: EventConfiguration) =>
            event.options.subscriptionId === incomingEvent.subscriptionId,
        )
        for (const event of eventsUsingMicrosoftEvent) {
          await this.app.handleEvent(event.id, {
            ...event,
            ...req.body,
          })
        }
      }
    }
    return true
  }

  sdk(): Client {
    return this.client
  }
}
