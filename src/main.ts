import { IUpdateInfo, updateElectronApp } from "update-electron-app";

import { execFileSync } from "node:child_process";
import { rmSync } from "node:fs";

import { BrowserWindow, Notification, app, shell } from "electron";
import started from "electron-squirrel-startup";

import { initAutoLaunch } from "./native/autoLaunch";
import { config } from "./native/config";
import { initDiscordRpc } from "./native/discordRpc";
import { initTray } from "./native/tray";
import { initVirtualMic } from "./native/virtualMic";
import { BUILD_URL, createMainWindow, mainWindow } from "./native/window";

// Desinstalar precisa apagar tambem a pasta de dados (Callju, dentro do AppData).
// O Squirrel so remove a pasta do programa, e o que sobrava ali, principalmente
// a copia do site guardada pelo service worker, fazia a reinstalacao abrir
// travada do mesmo jeito que antes.
if (
  process.platform === "win32" &&
  process.argv.includes("--squirrel-uninstall")
) {
  try {
    rmSync(app.getPath("userData"), { recursive: true, force: true });
  } catch {
    // Arquivo preso por outro processo: o resto sai mesmo assim
  }
}

// Squirrel-specific logic
// create/remove shortcuts on Windows when installing / uninstalling
// we just need to close out of the app immediately
if (started) {
  app.quit();
}

// disable hw-accel if so requested
if (!config.hardwareAcceleration) {
  app.disableHardwareAcceleration();
}

// ensure only one copy of the application can run
const acquiredLock = app.requestSingleInstanceLock();

// Primeira abertura depois de instalar, com outra copia ainda rodando.
//
// Fechar a janela so esconde o app na bandeja, entao a versao antiga quase
// sempre continua viva. Sem isto a versao nova, ao abrir, so trazia pra frente
// a janela da antiga, travada, e parecia que instalar nao tinha adiantado
// nada. Aqui a antiga e encerrada e a nova abre no lugar dela.
if (
  !acquiredLock &&
  process.platform === "win32" &&
  process.argv.includes("--squirrel-firstrun")
) {
  try {
    execFileSync("taskkill", [
      "/F",
      "/IM",
      "callju.exe",
      "/FI",
      `PID ne ${process.pid}`,
    ]);
  } catch {
    // Nada pra encerrar
  }
  app.relaunch({
    args: process.argv.slice(1).filter((a) => a !== "--squirrel-firstrun"),
  });
  app.exit(0);
}

const onNotifyUser = (_info: IUpdateInfo) => {
  const notification = new Notification({
    title: "Atualização do Callju pronta",
    body: "Pra instalar, clique com o botão direito no ícone do Callju perto do relógio, escolha Sair e abra de novo.",
    silent: true,
  });

  notification.show();
};

if (acquiredLock) {
  // start auto update logic
  updateElectronApp({ onNotifyUser });

  // create and configure the app when electron is ready
  app.on("ready", () => {
    // create window and application contexts
    createMainWindow();

    // save first launch state
    if (config.firstLaunch) {
      // Doesn't do anything right now. Used to enable auto start, but that behaviour was removed.
      // Left in case it gets used in the future.
      config.firstLaunch = false;
    }

    initTray();
    initDiscordRpc();
    initVirtualMic();
    initAutoLaunch();

    // Windows specific fix for notifications
    if (process.platform === "win32") {
      app.setAppUserModelId("chat.stoat.notifications");
    }
  });

  // focus the window if we try to launch again
  app.on("second-instance", () => {
    mainWindow.show();
    mainWindow.restore();
    mainWindow.focus();
  });

  // macOS specific behaviour to keep app active in dock:
  // (irrespective of the minimise-to-tray option)

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
      app.quit();
    }
  });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    } else {
      mainWindow.show();
      mainWindow.focus();
    }
  });

  // ensure URLs launch in external context
  app.on("web-contents-created", (_, contents) => {
    // prevent navigation out of build URL origin
    contents.on("will-navigate", (event, navigationUrl) => {
      if (new URL(navigationUrl).origin !== BUILD_URL.origin) {
        event.preventDefault();

        // Antes o clique simplesmente nao fazia nada, e parecia botao
        // quebrado: sair do Callju era proibido e ponto. Agora o endereco vai
        // pro navegador do sistema, que e onde ele deve abrir mesmo.
        if (/^https?:/.test(navigationUrl)) {
          setImmediate(() => shell.openExternal(navigationUrl));
        }
      }
    });

    // handle links externally
    contents.setWindowOpenHandler(({ url }) => {
      if (
        url.startsWith("http:") ||
        url.startsWith("https:") ||
        url.startsWith("mailto:")
      ) {
        setImmediate(() => {
          shell.openExternal(url);
        });
      }

      return { action: "deny" };
    });
  });
} else {
  app.quit();
}
