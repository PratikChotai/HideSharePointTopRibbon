import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";


import { PermissionKind } from "@pnp/sp/security";

const LOG_SOURCE: string = 'HideItApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHideItApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HideItApplicationCustomizer extends BaseApplicationCustomizer<IHideItApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    // in SPFx only
    sp.setup(this.context);
    try {
      let perms = await sp.web.getCurrentUserEffectivePermissions();
      if (sp.web.hasPermissions(perms, PermissionKind.UpdatePersonalWebParts)) {
        Log.info(LOG_SOURCE, `Should not hide ribbon for admin user.`);
        var odSuiteNav = document.getElementsByClassName("od-SuiteNav");
        for ( var i = 0; i < odSuiteNav.length; i++ ) {
          odSuiteNav[i]["style"].display = 'none';
        }

        var SuiteNavWrapper = document.getElementById("SuiteNavWrapper");
        SuiteNavWrapper.style.display = 'none';

      }
      else {
        Log.info(LOG_SOURCE, `Hide top ribbon for non admin user.`);

        var odSuiteNav = document.getElementsByClassName("od-SuiteNav");
        for ( var i = 0; i < odSuiteNav.length; i++ ) {
          odSuiteNav[i]["style"].display = 'none';
        }

        var SuiteNavWrapper = document.getElementById("SuiteNavWrapper");
        SuiteNavWrapper.style.display = 'none';
      }
      return Promise.resolve();
    }
    catch (error) {
      console.error(error);
      return Promise.resolve();
    }
  }
}
