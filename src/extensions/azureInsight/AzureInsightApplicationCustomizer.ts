import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { AppInsights} from 'applicationinsights-js'

import * as strings from 'AzureInsightApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AzureInsightApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAzureInsightApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AzureInsightApplicationCustomizer
  extends BaseApplicationCustomizer<IAzureInsightApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

     /* update with YOUR App Insights key: */
     let appInsightsKey: string = "51384f5f-c354-44af-accc-d4f5c6b05c8e";

     AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });
 
     // simple usage - all params will be derived..
     AppInsights.trackPageView();
 
     console.log(`OnInit: Called trackPageView().`);

    Dialog.alert(`Helloo from ${strings.Title}:\n\n${message}`);

    return Promise.resolve();
  
  }

  @override
  public onRender(): void {
  
    
  }
}
