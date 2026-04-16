import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI | undefined = undefined;
let _spContext: WebPartContext | undefined = undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context && (!_sp || _spContext !== context)) {
    _spContext = context;
    // Initializing with SPFx context
    _sp = spfi().using(SPFx(context));
  }
  return _sp as SPFI;
};

export const getSPContext = (): WebPartContext | undefined => {
  return _spContext;
};
