declare global {
  interface Window {
    electronAPI: {
      saveFileDialog: (options: { defaultPath: string }) => Promise<string | null>;
      selectDirectory: () => Promise<string | null>;
    };
  }
}

export {};
