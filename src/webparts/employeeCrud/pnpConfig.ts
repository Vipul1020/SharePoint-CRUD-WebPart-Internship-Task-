import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const initSP = (context: WebPartContext): void => {
  sp.setup({ spfxContext: context as any });
};

export { sp };
