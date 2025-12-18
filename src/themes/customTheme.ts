import { createLightTheme, createDarkTheme, type BrandVariants } from '@fluentui/react-components';

// SharePoint green (#217346) brand variants
const greenBrand: BrandVariants = {
  10: '#0a1f12',
  20: '#0e3219',
  30: '#124420',
  40: '#165727',
  50: '#1a6a2e',
  60: '#1e7d35',
  70: '#217346', // Primary green
  80: '#3d8a5a',
  90: '#5aa16f',
  100: '#77b884',
  110: '#94cf99',
  120: '#b1e6ae',
  130: '#cef8c3',
  140: '#e8fce2',
  150: '#f5fef3',
  160: '#fafffa',
};

export const greenLightTheme = createLightTheme(greenBrand);
export const greenDarkTheme = createDarkTheme(greenBrand);
