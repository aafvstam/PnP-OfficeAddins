import { getGlobal } from "../src/commands/commands";

export const SetRuntimeVisibleHelper = (visible: boolean) => {
  let p: any;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  } else {
    p = Office.addin.hide();
  }

  return p
    .then(() => {
      return visible;
    })
    .catch(error => {
      return error.code;
    });
};

export function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal() as any;

  // @ts-ignore
  OfficeRuntime.ui
    .getRibbon()
    // @ts-ignore
    .then(ribbon => {
      ribbon.requestUpdate({
        tabs: [
          {
            id: "ShareTime",
            controls: [
              {
                id: "BtnOpenTaskpane",
                enabled: !g.state.isTaskpaneOpen
              },
              {
                id: "BtnCloseTaskpane",
                enabled: g.state.isTaskpaneOpen
              }
            ]
          }
        ]
      });
    });
}

// This will check if state is initialized, and if not, initialize it.
// Useful as there are multiple entry points that need the state and it is not clear which one will get called first.
export async function ensureStateInitialized(isOfficeInitializing: boolean) {
  console.log("ensureInitialize called");
  let g = getGlobal() as any;
  let initValue = false;
  if (isOfficeInitializing) {
    //we are being called in response to Office Initialize
    if (g.state !== undefined) {
      if (g.state.isInitialized === false) {
        g.state.isInitialized = true;
      }
    }
    if (g.state === undefined) {
      initValue = true;
    }
  }

  if (g.state === undefined) {
    g.state = {
      isTaskpaneOpen: false,
      isSumEnabled: false,
      isInitialized: initValue,
      setTaskpaneStatus: (opened: boolean) => {
        g.state.isTaskpaneOpen = opened;
        updateRibbon();
      }
    };
    // console.log("init value:" + initValue);
    // console.log("is initialized: " + g.state.isInitialized);
    // monitorSheetChanges();
  }
  updateRibbon();
}
