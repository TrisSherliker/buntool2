/**
 * Configuration class for BunTool generator. Takes options from frontend
 * for parsing during bundle processing. 
 */

export const validFonts = [
    "serif",
    "traditional",
    "sansSerif",
    "monospaced",
    "times", 
    "helvetica", 
    "courier"
  ];

export const fontDisplayNames = {
  "serif": "Serif (Noto Serif)",
  "traditional": "Traditional (EB Garamond)",
  "sansSerif": "Sans Serif (Plus Jakarta)",
  "monospaced": "Monospaced (Ubuntu Mono)",
  "times": "Times New Roman",
  "helvetica": "Helvetica",
  "courier": "Courier"
};

export const validAlignments = [
    "left", 
    "centre", 
    "center",
    "right"
  ];

export const validNumberingStyles = [
    "PageX",
    "PageXofY",
    "X",
    "XslashY",
    "XofY",
    "None",
  ];

export const numberingStyleDisplayNames = {
    "PageX": "Page X",
    "PageXofY": "Page X of Y",
    "X": "X",
    "XslashY": "X/Y",
    "XofY": "X of Y",
    "None": "None",
};

export const validDateStyles = [
    "YYYY-MM-DD",
    "DD-MM-YYYY",
    "MM/DD/YYYY",
    "DD Mon. YYYY",
    "DD Month YYYY",
    "Mon DD, YYYY",
    "Month DD, YYYY",
    "None",
  ];

export const validOutlineStyles = [
    "plain", 
    "withPage", 
    "withDate", 
    "withDateandPage",
  ];

export const outlineStyleDisplayNames = {
    "plain": "Document title",
    "withPage": "Document title [page no]",
    "withDate": "Document title [date]",
    "withDateandPage": "Document title (date) [page no]",
};

export const validPrintableBundle = [
    true,
    false,
  ];

export const justTheIndex = [
    true,
    false,
  ];

class Config {
  /* 
  * Initialise with default configuration
  */
  constructor() {
    this.options = {
      heading: {
        claimNumber: "", // Default: blank
        bundleTitle: "Bundle", // Default: "Bundle"
        projectName: "", // Default: blank
        confidential: false, // Default: false
      },
      pageNumbering: {
        footerFont: "helvetica", // Default: helvetica
        alignment: "right", // Default: Right
        numberingStyle: "PageX", // Default: Page [X]
        footerPrefix: "", // Default: blank
      },
      index: {
        fontFace: "helvetica", // Default: helvetica
        dateStyle: "YYYY-MM-DD", // Default: YYYY-MM-DD
        outlineItemStyle: "withPage", // Default: with page
        justTheIndex: false, // Default: false
      },
      pageOptions: {
        printableBundle: false, // Default: false
      }
    };
  }
  
  /**
   * Method to update options
   * Options mainly passed in from frontend.
   */
    updateOptions(newOptions) {
      this.options = {
        heading: { ...this.options.heading, ...newOptions.heading },
        pageNumbering: { ...this.options.pageNumbering, ...newOptions.pageNumbering },
        index: { ...this.options.index, ...newOptions.index },
        pageOptions: { ...this.options.pageOptions, ...newOptions.pageOptions },
      };
    }

  /**
   * Method to validate options
   * Defines valid values
   * Errors thrown for invalid options
  */
  validateOptions() {
    if (!validFonts.includes(this.options.index.fontFace)) {
      throw new Error(`Invalid index font: ${this.options.index.fontFace}`);
    }
    if (!validFonts.includes(this.options.pageNumbering.footerFont)) {
        throw new Error(`Invalid footer font: ${this.options.pageNumbering.footerFont}`);
    }
    if (!validAlignments.includes(this.options.pageNumbering.alignment)) {
      throw new Error(`Invalid alignment: ${this.options.pageNumbering.alignment}`);
    }
    if (!validNumberingStyles.includes(this.options.pageNumbering.numberingStyle)) {
      throw new Error(`Invalid numbering style: ${this.options.pageNumbering.numberingStyle}`);
    }
    if (!validDateStyles.includes(this.options.index.dateStyle)) {
      throw new Error(`Invalid date style: ${this.options.index.dateStyle}`);
    }
    if (!validOutlineStyles.includes(this.options.index.outlineItemStyle)) {
      throw new Error(`Invalid outline item style: ${this.options.index.outlineItemStyle}`);
    }
    if (!validPrintableBundle.includes(this.options.pageOptions.printableBundle)) {
      throw new Error(`Invalid printable bundle option: ${this.options.pageOptions.printableBundle}`);
    }

    if (!justTheIndex.includes(this.options.index.justTheIndex)) {
      throw new Error(`Invalid justTheIndex option: ${this.options.index.justTheIndex}`);
    }
  }
      
  /**
   * Method to validate structure
   * Defines required fields
   * Errors thrown for missing fields
  */  
 validateStructure() {
    const requiredPaths = {
      heading: ["claimNumber", "bundleTitle", "projectName", "confidential"],
      pageNumbering: ["footerFont", "alignment", "numberingStyle", "footerPrefix"],
      index: ["fontFace", "dateStyle", "outlineItemStyle", "justTheIndex"],
      pageOptions: ["printableBundle"],
    };
    
    for (const [section, fields] of Object.entries(requiredPaths)) {
      if (!this.options[section]) {
        throw new Error(`Invalid config: Missing configuration section: ${section}`);
      }
      for (const field of fields) {
        if (this.options[section][field] === undefined) {
          throw new Error(`Invalid config: Missing configuration field: ${section}.${field}`);
        }
      }
    }
  }

  /**
   * Method to retrieve option by key path
   * returns value, or null if not found
   */
  getOption(key) {
    return key.split(".").reduce((obj, k) => (obj && obj[k] !== undefined ? obj[k] : null), this.options);
    }
  }

export default Config;