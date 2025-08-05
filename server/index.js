#!/usr/bin/env node

const { execFile } = require('child_process');
const { promisify } = require('util');

const execFileAsync = promisify(execFile);

console.error("Safari AppleScript MCP server starting...");

// Constants
const APPLESCRIPT_TIMEOUT = 10000; // 10 seconds
const MAX_RETRIES = 3;
const RETRY_DELAY = 1000; // 1 second

// Universal type caster for AppleScript values
function universalCast(value) {
  if (value === null || value === undefined) return null;
  
  const str = String(value).trim();
  if (str === '') return '';
  
  // Boolean
  const lower = str.toLowerCase();
  if (['true', 'yes'].includes(lower)) return true;
  if (['false', 'no'].includes(lower)) return false;
  
  // Number
  if (/^-?\d+(\.\d+)?$/.test(str)) {
    const num = Number(str);
    return Number.isInteger(num) ? num : num;
  }
  
  // List/Record - anything in braces, return as-is for AppleScript
  if (str.startsWith('{') && str.endsWith('}')) {
    return str; // AppleScript will handle the parsing
  }
  
  // Auto-detect comma-separated values that should be lists/rectangles
  if (str.includes(',') && !str.startsWith('{')) {
    const parts = str.split(',').map(p => p.trim());
    
    // Rectangle pattern: 4 numbers (x, y, width, height)
    if (parts.length === 4 && parts.every(p => /^-?\d+(\.\d+)?$/.test(p))) {
      return `{${str}}`; // Add brackets for rectangle
    }
    
    // Generic list: 2+ comma-separated values
    if (parts.length >= 2) {
      return `{${str}}`; // Add brackets for list
    }
  }
  
  // Date patterns
  if (str.startsWith('date "') || /^\d{4}-\d{2}-\d{2}/.test(str)) {
    return `date "${str}"`;
  }
  
  // String - remove quotes if present
  if ((str.startsWith('"') && str.endsWith('"')) || 
      (str.startsWith("'") && str.endsWith("'"))) {
    return str.slice(1, -1);
  }
  
  return str;
}

// Cast and escape for AppleScript injection
function castAndEscape(value) {
  const casted = universalCast(value);
  
  // If it's a string that doesn't start with {, escape it
  if (typeof casted === 'string' && !casted.startsWith('{') && !casted.startsWith('date')) {
    return escapeForAppleScript(casted);
  }
  
  // Numbers, booleans, and AppleScript literals go as-is
  return casted;
}

// Helper function to escape strings for AppleScript
function escapeForAppleScript(str) {
  if (typeof str !== "string") return str;
  return str
    .replace(/\\/g, "\\\\") // Escape backslashes first
    .replace(/"/g, '\"') // Then escape double quotes
    .replace(/\n/g, "\\n") // Escape newlines
    .replace(/\r/g, "\\r"); // Escape carriage returns
}

// Test if Safari is available
async function checkSafariAvailable() {
  try {
    const script = 'tell application "Safari" to return "available"';
    const result = await executeAppleScript(script);
    return result === "available";
  } catch (error) {
    return false;
  }
}

// Execute AppleScript with retry logic
async function executeAppleScript(script, retries = MAX_RETRIES) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const { stdout, stderr } = await execFileAsync(
        "osascript",
        ["-e", script],
        {
          timeout: APPLESCRIPT_TIMEOUT,
          maxBuffer: 1024 * 1024, // 1MB buffer
        },
      );
      if (stderr) {
        console.error("AppleScript stderr:", stderr);
      }
      return stdout.trim();
    } catch (error) {
      if (attempt === retries) {
        console.error("AppleScript execution error after retries:", error);
        throw new Error(`AppleScript error: ${error.message}`);
      }
      await new Promise((resolve) =>
        setTimeout(resolve, RETRY_DELAY * Math.pow(2, attempt)),
      );
    }
  }
}

// MCP server implementation
class SafariMCPServer {
  constructor() {
    this.initialized = false;
    this.setupStdio();
  }

  setupStdio() {
    process.stdin.setEncoding('utf8');
    
    let buffer = '';
    process.stdin.on('data', (data) => {
      buffer += data;
      const lines = buffer.split('\n');
      buffer = lines.pop() || ''; // Keep incomplete line in buffer
      
      lines.forEach(line => {
        if (line.trim()) {
          this.handleMessage(line.trim());
        }
      });
    });
  }

  async handleMessage(data) {
    try {
      const request = JSON.parse(data);
      console.error("Received request:", request.method, request.id);
      
      if (request.method === 'initialize') {
        await this.handleInitialize(request);
      } else if (request.method === 'initialized') {
        await this.handleInitialized(request);
      } else if (request.method === 'tools/list') {
        await this.handleToolsList(request);
      } else if (request.method === 'tools/call') {
        await this.handleToolsCall(request);
      } else {
        console.error("Unknown method:", request.method);
      }
    } catch (error) {
      console.error("Error processing message:", error);
    }
  }

  async handleInitialize(request) {
    console.error("Handling initialize request");
    const response = {
      jsonrpc: '2.0',
      id: request.id,
      result: {
        protocolVersion: '2024-11-05',
        capabilities: {
          tools: {}
        },
        serverInfo: {
          name: 'safari-applescript',
          version: '0.1.0'
        }
      }
    };
    this.sendResponse(response);
  }

  async handleInitialized(request) {
    console.error("Handling initialized notification");
    this.initialized = true;
  }

  async handleToolsList(request) {
    console.error("Handling tools/list request");
    const response = {
      jsonrpc: '2.0',
      id: request.id,
      result: {
        tools: [
          {
  name: 'open',
  description: 'Open a document.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_file: {
        type: 'string',
        description: 'The file(s) to be opened.'
      }
    },
    required: ['direct_parameter_required_file'],
    additionalProperties: false
  }
},
          {
  name: 'close_document',
  description: 'Close a document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      saving_optional_save_options: {
        type: 'string',
        description: 'Should changes be saved before closing?'
      },
      saving_in_optional_file: {
        type: 'string',
        description: 'The file in which to save the document, if so.'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'close_tab_of_window',
  description: 'Close a tab of window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      saving_optional_save_options: {
        type: 'string',
        description: 'Should changes be saved before closing?'
      },
      saving_in_optional_file: {
        type: 'string',
        description: 'The file in which to save the document, if so.'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'close_window',
  description: 'Close a window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      saving_optional_save_options: {
        type: 'string',
        description: 'Should changes be saved before closing?'
      },
      saving_in_optional_file: {
        type: 'string',
        description: 'The file in which to save the document, if so.'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'save_document',
  description: 'Save a document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      inParam_optional_file: {
        type: 'string',
        description: 'The file in which to save the document.'
      },
      as_optional_saveable_file_format: {
        type: 'string',
        description: 'The file format to use.'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'save_window',
  description: 'Save a window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      inParam_optional_file: {
        type: 'string',
        description: 'The file in which to save the document.'
      },
      as_optional_saveable_file_format: {
        type: 'string',
        description: 'The file format to use.'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'print_document',
  description: 'Print a document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      with_properties_optional_print_settings: {
        type: 'string',
        description: 'The print settings to use.'
      },
      print_dialog_optional_boolean: {
        type: 'boolean',
        description: 'Should the application show the print dialog?'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'print_window',
  description: 'Print a window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      with_properties_optional_print_settings: {
        type: 'string',
        description: 'The print settings to use.'
      },
      print_dialog_optional_boolean: {
        type: 'boolean',
        description: 'Should the application show the print dialog?'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'quit',
  description: 'Quit the application.',
  inputSchema: {
    type: 'object',
    properties: {
      saving_optional_save_options: {
        type: 'string',
        description: 'Should changes be saved before quitting?'
      }
    },
    additionalProperties: false
  }
},
          {
  name: 'count_document',
  description: 'Return the number of elements of a particular class within a document.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'count_tab_of_window',
  description: 'Return the number of elements of a particular class within a tab of window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'count_window',
  description: 'Return the number of elements of a particular class within a window.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'delete',
  description: 'Delete an object.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_specifier: {
        type: 'string',
        description: 'The object(s) to delete.'
      }
    },
    required: ['direct_parameter_required_specifier'],
    additionalProperties: false
  }
},
          {
  name: 'duplicate',
  description: 'Copy an object.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_specifier: {
        type: 'string',
        description: 'The object(s) to copy.'
      },
      to_optional_location_specifier: {
        type: 'string',
        description: 'The location for the new copy or copies.'
      },
      with_properties_optional_record: {
        type: 'string',
        description: 'Properties to set in the new copy or copies right away.'
      }
    },
    required: ['direct_parameter_required_specifier'],
    additionalProperties: false
  }
},
          {
  name: 'exists',
  description: 'Verify that an object exists.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_any: {
        type: 'string',
        description: 'The object(s) to check.'
      }
    },
    required: ['direct_parameter_required_any'],
    additionalProperties: false
  }
},
          {
  name: 'make_document',
  description: 'Create a new document.',
  inputSchema: {
    type: 'object',
    properties: {
      at_optional_location_specifier: {
        type: 'string',
        description: 'The location at which to insert the object.'
      },
      with_data_optional_any: {
        type: 'string',
        description: 'The initial contents of the object.'
      },
      with_properties_optional_text_url: {
        type: 'string',
        description: 'Optional URL property: The current URL of the document.'
      }
    },
    additionalProperties: false
  }
},
          {
  name: 'make_tab_of_window',
  description: 'Create a new tab of window.',
  inputSchema: {
    type: 'object',
    properties: {
      at_required_location_specifier_window: {
        type: 'string',
        description: 'The window location where the tab should be created (e.g., \"window 1\")'
      },
      with_data_optional_any: {
        type: 'string',
        description: 'The initial contents of the object.'
      },
      with_properties_optional_text_url: {
        type: 'string',
        description: 'Optional URL property: The current URL of the tab.'
      }
    },
    required: ['at_required_location_specifier_window'],
    additionalProperties: false
  }
},
          {
  name: 'make_window',
  description: 'Create a new window.',
  inputSchema: {
    type: 'object',
    properties: {
      at_optional_location_specifier: {
        type: 'string',
        description: 'The location at which to insert the object.'
      },
      with_data_optional_any: {
        type: 'string',
        description: 'The initial contents of the object.'
      },
      with_properties_optional_integer_index: {
        type: 'number',
        description: 'Optional index property: The index of the window, ordered front to back.'
      },
      with_properties_optional_tab_current_tab: {
        type: 'string',
        description: 'Optional current tab property: The current tab.'
      },
      with_properties_optional_boolean_zoomed: {
        type: 'boolean',
        description: 'Optional zoomed property: Is the window zoomed right now?'
      },
      with_properties_optional_boolean_miniaturized: {
        type: 'boolean',
        description: 'Optional miniaturized property: Is the window minimized right now?'
      },
      with_properties_optional_boolean_visible: {
        type: 'boolean',
        description: 'Optional visible property: Is the window visible right now?'
      },
      with_properties_optional_rectangle_bounds: {
        type: 'string',
        description: 'Optional bounds property: The bounding rectangle of the window.'
      }
    },
    additionalProperties: false
  }
},
          {
  name: 'move',
  description: 'Move an object to a new location.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_specifier: {
        type: 'string',
        description: 'The object(s) to move.'
      },
      to_required_location_specifier: {
        type: 'string',
        description: 'The new location for the object(s).'
      }
    },
    required: ['direct_parameter_required_specifier', 'to_required_location_specifier'],
    additionalProperties: false
  }
},
          {
  name: 'get_name_of_application',
  description: 'Get The name of the application.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'get_frontmost_of_application',
  description: 'Get Is this the active application?',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'get_version_of_application',
  description: 'Get The version number of the application.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'get_name_of_document',
  description: 'Get Its name. of document',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_modified_of_document',
  description: 'Get Has it been modified since the last save? of document',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_file_of_document',
  description: 'Get Its location on disk, if it has one. of document',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_name_of_window',
  description: 'Get The title of the window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_id_of_window',
  description: 'Get The unique identifier of the window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_index_of_window',
  description: 'Get The index of the window, ordered front to back.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_index_of_window',
  description: 'Set The index of the window, ordered front to back.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_integer: {
        type: 'number',
        description: 'New value for The index of the window, ordered front to back.'
      }
    },
    required: ['target_window_required_string', 'value_required_integer'],
    additionalProperties: false
  }
},
          {
  name: 'get_bounds_of_window',
  description: 'Get The bounding rectangle of the window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_bounds_of_window',
  description: 'Set The bounding rectangle of the window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_rectangle: {
        type: 'string',
        description: 'New value for The bounding rectangle of the window.'
      }
    },
    required: ['target_window_required_string', 'value_required_rectangle'],
    additionalProperties: false
  }
},
          {
  name: 'get_closeable_of_window',
  description: 'Get Does the window have a close button?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_miniaturizable_of_window',
  description: 'Get Does the window have a minimize button?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_miniaturized_of_window',
  description: 'Get Is the window minimized right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_miniaturized_of_window',
  description: 'Set Is the window minimized right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_boolean: {
        type: 'boolean',
        description: 'New value for Is the window minimized right now?'
      }
    },
    required: ['target_window_required_string', 'value_required_boolean'],
    additionalProperties: false
  }
},
          {
  name: 'get_resizable_of_window',
  description: 'Get Can the window be resized?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_visible_of_window',
  description: 'Get Is the window visible right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_visible_of_window',
  description: 'Set Is the window visible right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_boolean: {
        type: 'boolean',
        description: 'New value for Is the window visible right now?'
      }
    },
    required: ['target_window_required_string', 'value_required_boolean'],
    additionalProperties: false
  }
},
          {
  name: 'get_zoomable_of_window',
  description: 'Get Does the window have a zoom button?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_zoomed_of_window',
  description: 'Get Is the window zoomed right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_zoomed_of_window',
  description: 'Set Is the window zoomed right now?',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_boolean: {
        type: 'boolean',
        description: 'New value for Is the window zoomed right now?'
      }
    },
    required: ['target_window_required_string', 'value_required_boolean'],
    additionalProperties: false
  }
},
          {
  name: 'get_document_of_window',
  description: 'Get The document whose contents are displayed in the window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'add_reading_list_item',
  description: 'Add a new Reading List item with the given URL. Allows a custom title and preview text to be specified.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_text: {
        type: 'string',
        description: 'URL of the Reading List item'
      },
      and_preview_text_optional_text: {
        type: 'string',
        description: 'Preview text for the Reading List item, usually the first few sentences of the article'
      },
      with_title_optional_text: {
        type: 'string',
        description: 'Title of the Reading List item'
      }
    },
    required: ['direct_parameter_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'do_javascript_document',
  description: 'Applies a string of JavaScript code to a document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      inParam_optional_document: {
        type: 'string',
        description: 'The tab that the JavaScript should be evaluated in.'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'do_javascript_tab_of_window',
  description: 'Applies a string of JavaScript code to a tab of window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      inParam_optional_document: {
        type: 'string',
        description: 'The tab that the JavaScript should be evaluated in.'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'search_the_web_document',
  description: 'Searches the web using Safari\'s current search provider for document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      inParam_optional_document: {
        type: 'string',
        description: 'The tab that the search results should shown in.'
      },
      forParam_required_text: {
        type: 'string',
        description: 'The query to search for.'
      }
    },
    required: ['target_document_required_string', 'forParam_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'search_the_web_tab_of_window',
  description: 'Searches the web using Safari\'s current search provider for tab of window.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      inParam_optional_document: {
        type: 'string',
        description: 'The tab that the search results should shown in.'
      },
      forParam_required_text: {
        type: 'string',
        description: 'The query to search for.'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string', 'forParam_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'show_bookmarks',
  description: 'Shows Safari\'s bookmarks.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'show_extensions_preferences',
  description: 'Show Safari Extensions preferences.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_text: {
        type: 'string',
        description: 'The identifier of the extension to select.'
      }
    },
    required: ['direct_parameter_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'dispatch_message_to_extension',
  description: 'Dispatch a message to a Safari Extension.',
  inputSchema: {
    type: 'object',
    properties: {
      direct_parameter_required_any: {
        type: 'string',
        description: 'A dictionary describing the message'
      }
    },
    required: ['direct_parameter_required_any'],
    additionalProperties: false
  }
},
          {
  name: 'sync_all_plist_to_disk',
  description: 'Make sure that all in-memory structures are in-sync with their on-disk counterparts.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'show_privacy_report',
  description: 'Show Safari\'s Privacy Report',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'show_credit_card_settings',
  description: 'Show Safari Credit Card Settings.',
  inputSchema: {
    type: 'object',
    properties: {
    },
    additionalProperties: false
  }
},
          {
  name: 'get_current_tab_of_window',
  description: 'Get The current tab. of window',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_current_tab_of_window',
  description: 'Set The current tab. of window',
  inputSchema: {
    type: 'object',
    properties: {
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_tab: {
        type: 'string',
        description: 'New value for The current tab.'
      }
    },
    required: ['target_window_required_string', 'value_required_tab'],
    additionalProperties: false
  }
},
          {
  name: 'get_source_of_document',
  description: 'Get The HTML source of the web page currently loaded in the document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_url_of_document',
  description: 'Get The current URL of the document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_url_of_document',
  description: 'Set The current URL of the document.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      },
      value_required_text: {
        type: 'string',
        description: 'New value for The current URL of the document.'
      }
    },
    required: ['target_document_required_string', 'value_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'get_text_of_document',
  description: 'Get The text of the web page currently loaded in the document. Modifications to text aren\'t reflected on the web page.',
  inputSchema: {
    type: 'object',
    properties: {
      target_document_required_string: {
        type: 'string',
        description: 'The document object to access (e.g., \"front document\", \"document 1\")'
      }
    },
    required: ['target_document_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_source_of_tab_of_window',
  description: 'Get The HTML source of the web page currently loaded in the tab.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_url_of_tab_of_window',
  description: 'Get The current URL of the tab.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'set_url_of_tab_of_window',
  description: 'Set The current URL of the tab.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      },
      value_required_text: {
        type: 'string',
        description: 'New value for The current URL of the tab.'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string', 'value_required_text'],
    additionalProperties: false
  }
},
          {
  name: 'get_index_of_tab_of_window',
  description: 'Get The index of the tab, ordered left to right.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_text_of_tab_of_window',
  description: 'Get The text of the web page currently loaded in the tab. Modifications to text aren\'t reflected on the web page.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_visible_of_tab_of_window',
  description: 'Get Whether the tab is currently visible.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_name_of_tab_of_window',
  description: 'Get The name of the tab.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
},
          {
  name: 'get_pid_of_tab_of_window',
  description: 'Get The pid of the WebContent process backing the tab, if it exists.',
  inputSchema: {
    type: 'object',
    properties: {
      target_tab_required_string: {
        type: 'string',
        description: 'The tab object to access (e.g., \"front tab\", \"tab 1\")'
      },
      target_window_required_string: {
        type: 'string',
        description: 'The window object to access (e.g., \"front window\", \"window 1\")'
      }
    },
    required: ['target_tab_required_string', 'target_window_required_string'],
    additionalProperties: false
  }
}
        ]
      }
    };
    this.sendResponse(response);
  }

  async handleToolsCall(request) {
    console.error("Handling tools/call request for:", request.params.name);
    
    try {
      // Check app availability first (skip for static data functions)
      const staticFunctions = ['get_all_classes', 'get_all_properties_of', 'get_parsed_sdef'];
      if (!staticFunctions.includes(request.params.name)) {
        const isSafariAvailable = await checkSafariAvailable();
        if (!isSafariAvailable) {
          const errorResponse = {
            jsonrpc: '2.0',
            id: request.id,
            result: {
              content: [{
                type: 'text',
                text: JSON.stringify({
                  success: false,
                  error: 'Application is not available or not running'
                }, null, 2)
              }]
            }
          };
          this.sendResponse(errorResponse);
          return;
        }
      }

      const { name, arguments: args } = request.params;
      let result;

      switch (name) {
        case 'open':
  result = await this.open(args.direct_parameter_required_file);
  break;
        case 'close_document':
  result = await this.closeDocument(args.target_document_required_string, args.saving_optional_save_options, args.saving_in_optional_file);
  break;
        case 'close_tab_of_window':
  result = await this.closeTabOfWindow(args.target_tab_required_string, args.target_window_required_string, args.saving_optional_save_options, args.saving_in_optional_file);
  break;
        case 'close_window':
  result = await this.closeWindow(args.target_window_required_string, args.saving_optional_save_options, args.saving_in_optional_file);
  break;
        case 'save_document':
  result = await this.saveDocument(args.target_document_required_string, args.inParam_optional_file, args.as_optional_saveable_file_format);
  break;
        case 'save_window':
  result = await this.saveWindow(args.target_window_required_string, args.inParam_optional_file, args.as_optional_saveable_file_format);
  break;
        case 'print_document':
  result = await this.printDocument(args.target_document_required_string, args.with_properties_optional_print_settings, args.print_dialog_optional_boolean);
  break;
        case 'print_window':
  result = await this.printWindow(args.target_window_required_string, args.with_properties_optional_print_settings, args.print_dialog_optional_boolean);
  break;
        case 'quit':
  result = await this.quit(args.saving_optional_save_options);
  break;
        case 'count_document':
  result = await this.countDocument();
  break;
        case 'count_tab_of_window':
  result = await this.countTabOfWindow(args.target_window_required_string);
  break;
        case 'count_window':
  result = await this.countWindow();
  break;
        case 'delete':
  result = await this.delete(args.direct_parameter_required_specifier);
  break;
        case 'duplicate':
  result = await this.duplicate(args.direct_parameter_required_specifier, args.to_optional_location_specifier, args.with_properties_optional_record);
  break;
        case 'exists':
  result = await this.exists(args.direct_parameter_required_any);
  break;
        case 'make_document':
  result = await this.makeDocument(args.at_optional_location_specifier, args.with_data_optional_any, args.with_properties_optional_text_url);
  break;
        case 'make_tab_of_window':
  result = await this.makeTabOfWindow(args.at_required_location_specifier_window, args.with_data_optional_any, args.with_properties_optional_text_url);
  break;
        case 'make_window':
  result = await this.makeWindow(args.at_optional_location_specifier, args.with_data_optional_any, args.with_properties_optional_integer_index, args.with_properties_optional_tab_current_tab, args.with_properties_optional_boolean_zoomed, args.with_properties_optional_boolean_miniaturized, args.with_properties_optional_boolean_visible, args.with_properties_optional_rectangle_bounds);
  break;
        case 'move':
  result = await this.move(args.direct_parameter_required_specifier, args.to_required_location_specifier);
  break;
        case 'get_name_of_application':
  result = await this.getNameOfApplication();
  break;
        case 'get_frontmost_of_application':
  result = await this.getFrontmostOfApplication();
  break;
        case 'get_version_of_application':
  result = await this.getVersionOfApplication();
  break;
        case 'get_name_of_document':
  result = await this.getNameOfDocument(args.target_document_required_string);
  break;
        case 'get_modified_of_document':
  result = await this.getModifiedOfDocument(args.target_document_required_string);
  break;
        case 'get_file_of_document':
  result = await this.getFileOfDocument(args.target_document_required_string);
  break;
        case 'get_name_of_window':
  result = await this.getNameOfWindow(args.target_window_required_string);
  break;
        case 'get_id_of_window':
  result = await this.getIdOfWindow(args.target_window_required_string);
  break;
        case 'get_index_of_window':
  result = await this.getIndexOfWindow(args.target_window_required_string);
  break;
        case 'set_index_of_window':
  result = await this.setIndexOfWindow(args.target_window_required_string, args.value_required_integer);
  break;
        case 'get_bounds_of_window':
  result = await this.getBoundsOfWindow(args.target_window_required_string);
  break;
        case 'set_bounds_of_window':
  result = await this.setBoundsOfWindow(args.target_window_required_string, args.value_required_rectangle);
  break;
        case 'get_closeable_of_window':
  result = await this.getCloseableOfWindow(args.target_window_required_string);
  break;
        case 'get_miniaturizable_of_window':
  result = await this.getMiniaturizableOfWindow(args.target_window_required_string);
  break;
        case 'get_miniaturized_of_window':
  result = await this.getMiniaturizedOfWindow(args.target_window_required_string);
  break;
        case 'set_miniaturized_of_window':
  result = await this.setMiniaturizedOfWindow(args.target_window_required_string, args.value_required_boolean);
  break;
        case 'get_resizable_of_window':
  result = await this.getResizableOfWindow(args.target_window_required_string);
  break;
        case 'get_visible_of_window':
  result = await this.getVisibleOfWindow(args.target_window_required_string);
  break;
        case 'set_visible_of_window':
  result = await this.setVisibleOfWindow(args.target_window_required_string, args.value_required_boolean);
  break;
        case 'get_zoomable_of_window':
  result = await this.getZoomableOfWindow(args.target_window_required_string);
  break;
        case 'get_zoomed_of_window':
  result = await this.getZoomedOfWindow(args.target_window_required_string);
  break;
        case 'set_zoomed_of_window':
  result = await this.setZoomedOfWindow(args.target_window_required_string, args.value_required_boolean);
  break;
        case 'get_document_of_window':
  result = await this.getDocumentOfWindow(args.target_window_required_string);
  break;
        case 'add_reading_list_item':
  result = await this.addReadingListItem(args.direct_parameter_required_text, args.and_preview_text_optional_text, args.with_title_optional_text);
  break;
        case 'do_javascript_document':
  result = await this.doJavascriptDocument(args.target_document_required_string, args.inParam_optional_document);
  break;
        case 'do_javascript_tab_of_window':
  result = await this.doJavascriptTabOfWindow(args.target_tab_required_string, args.target_window_required_string, args.inParam_optional_document);
  break;
        case 'search_the_web_document':
  result = await this.searchTheWebDocument(args.target_document_required_string, args.inParam_optional_document, args.forParam_required_text);
  break;
        case 'search_the_web_tab_of_window':
  result = await this.searchTheWebTabOfWindow(args.target_tab_required_string, args.target_window_required_string, args.inParam_optional_document, args.forParam_required_text);
  break;
        case 'show_bookmarks':
  result = await this.showBookmarks();
  break;
        case 'show_extensions_preferences':
  result = await this.showExtensionsPreferences(args.direct_parameter_required_text);
  break;
        case 'dispatch_message_to_extension':
  result = await this.dispatchMessageToExtension(args.direct_parameter_required_any);
  break;
        case 'sync_all_plist_to_disk':
  result = await this.syncAllPlistToDisk();
  break;
        case 'show_privacy_report':
  result = await this.showPrivacyReport();
  break;
        case 'show_credit_card_settings':
  result = await this.showCreditCardSettings();
  break;
        case 'get_current_tab_of_window':
  result = await this.getCurrentTabOfWindow(args.target_window_required_string);
  break;
        case 'set_current_tab_of_window':
  result = await this.setCurrentTabOfWindow(args.target_window_required_string, args.value_required_tab);
  break;
        case 'get_source_of_document':
  result = await this.getSourceOfDocument(args.target_document_required_string);
  break;
        case 'get_url_of_document':
  result = await this.getUrlOfDocument(args.target_document_required_string);
  break;
        case 'set_url_of_document':
  result = await this.setUrlOfDocument(args.target_document_required_string, args.value_required_text);
  break;
        case 'get_text_of_document':
  result = await this.getTextOfDocument(args.target_document_required_string);
  break;
        case 'get_source_of_tab_of_window':
  result = await this.getSourceOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'get_url_of_tab_of_window':
  result = await this.getUrlOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'set_url_of_tab_of_window':
  result = await this.setUrlOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string, args.value_required_text);
  break;
        case 'get_index_of_tab_of_window':
  result = await this.getIndexOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'get_text_of_tab_of_window':
  result = await this.getTextOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'get_visible_of_tab_of_window':
  result = await this.getVisibleOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'get_name_of_tab_of_window':
  result = await this.getNameOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        case 'get_pid_of_tab_of_window':
  result = await this.getPidOfTabOfWindow(args.target_tab_required_string, args.target_window_required_string);
  break;
        default:
          throw new Error(`Unknown tool: ${name}`);
      }

      const response = {
        jsonrpc: '2.0',
        id: request.id,
        result: {
          content: [{
            type: 'text',
            text: JSON.stringify(result, null, 2)
          }]
        }
      };
      this.sendResponse(response);

    } catch (error) {
      console.error(`Error in tool '${request.params.name}':`, error);
      const errorResponse = {
        jsonrpc: '2.0',
        id: request.id,
        result: {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: false,
              error: error.message,
              tool: request.params.name,
              args: request.params.arguments
            }, null, 2)
          }]
        }
      };
      this.sendResponse(errorResponse);
    }
  }

  async open(direct_parameter_required_file) {
    if (direct_parameter_required_file === undefined || direct_parameter_required_file === null) {
      throw new Error("direct_parameter_required_file is required");
    }

    const castedDirect_parameter = direct_parameter_required_file ? castAndEscape(direct_parameter_required_file) : null;

    const script = `
      tell application "Safari"
        open ${castedDirect_parameter}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_file || null
    };
  }

  async closeDocument(target_document_required_string, saving_optional_save_options, saving_in_optional_file) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const castedDocument = castAndEscape(target_document_required_string);
    const castedSaving = saving_optional_save_options ? castAndEscape(saving_optional_save_options) : null;
    const valueForScriptSaving = castedSaving && typeof castedSaving === 'string' && !castedSaving.startsWith('{') && !castedSaving.startsWith('date') ? `"${castedSaving.replace(/"/g, "'")}"` : castedSaving;
    const castedSaving_in = saving_in_optional_file ? castAndEscape(saving_in_optional_file) : null;
    const valueForScriptSaving_in = castedSaving_in && typeof castedSaving_in === 'string' && !castedSaving_in.startsWith('{') && !castedSaving_in.startsWith('date') ? `"${castedSaving_in.replace(/"/g, "'")}"` : castedSaving_in;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        close ${castedDocument}${saving_optional_save_options ? ' saving ' + valueForScriptSaving : ''}${saving_in_optional_file ? ' saving in ' + valueForScriptSaving_in : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      document: target_document_required_string,
      saving: saving_optional_save_options || null,
      saving_in: saving_in_optional_file || null
    };
  }

  async closeTabOfWindow(target_tab_required_string, target_window_required_string, saving_optional_save_options, saving_in_optional_file) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedTab = castAndEscape(target_tab_required_string);
    const castedWindow = castAndEscape(target_window_required_string);
    const castedSaving = saving_optional_save_options ? castAndEscape(saving_optional_save_options) : null;
    const valueForScriptSaving = castedSaving && typeof castedSaving === 'string' && !castedSaving.startsWith('{') && !castedSaving.startsWith('date') ? `"${castedSaving.replace(/"/g, "'")}"` : castedSaving;
    const castedSaving_in = saving_in_optional_file ? castAndEscape(saving_in_optional_file) : null;
    const valueForScriptSaving_in = castedSaving_in && typeof castedSaving_in === 'string' && !castedSaving_in.startsWith('{') && !castedSaving_in.startsWith('date') ? `"${castedSaving_in.replace(/"/g, "'")}"` : castedSaving_in;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        close ${castedTab} of ${castedWindow}${saving_optional_save_options ? ' saving ' + valueForScriptSaving : ''}${saving_in_optional_file ? ' saving in ' + valueForScriptSaving_in : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string,
      saving: saving_optional_save_options || null,
      saving_in: saving_in_optional_file || null
    };
  }

  async closeWindow(target_window_required_string, saving_optional_save_options, saving_in_optional_file) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedWindow = castAndEscape(target_window_required_string);
    const castedSaving = saving_optional_save_options ? castAndEscape(saving_optional_save_options) : null;
    const valueForScriptSaving = castedSaving && typeof castedSaving === 'string' && !castedSaving.startsWith('{') && !castedSaving.startsWith('date') ? `"${castedSaving.replace(/"/g, "'")}"` : castedSaving;
    const castedSaving_in = saving_in_optional_file ? castAndEscape(saving_in_optional_file) : null;
    const valueForScriptSaving_in = castedSaving_in && typeof castedSaving_in === 'string' && !castedSaving_in.startsWith('{') && !castedSaving_in.startsWith('date') ? `"${castedSaving_in.replace(/"/g, "'")}"` : castedSaving_in;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        close ${castedWindow}${saving_optional_save_options ? ' saving ' + valueForScriptSaving : ''}${saving_in_optional_file ? ' saving in ' + valueForScriptSaving_in : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      window: target_window_required_string,
      saving: saving_optional_save_options || null,
      saving_in: saving_in_optional_file || null
    };
  }

  async saveDocument(target_document_required_string, inParam_optional_file, as_optional_saveable_file_format) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const castedDocument = castAndEscape(target_document_required_string);
    const castedIn = inParam_optional_file ? castAndEscape(inParam_optional_file) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;
    const castedAs = as_optional_saveable_file_format ? castAndEscape(as_optional_saveable_file_format) : null;
    const valueForScriptAs = castedAs && typeof castedAs === 'string' && !castedAs.startsWith('{') && !castedAs.startsWith('date') ? `"${castedAs.replace(/"/g, "'")}"` : castedAs;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        save ${castedDocument}${inParam_optional_file ? ' in ' + valueForScriptIn : ''}${as_optional_saveable_file_format ? ' as ' + valueForScriptAs : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      document: target_document_required_string,
      in: inParam_optional_file || null,
      as: as_optional_saveable_file_format || null
    };
  }

  async saveWindow(target_window_required_string, inParam_optional_file, as_optional_saveable_file_format) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedWindow = castAndEscape(target_window_required_string);
    const castedIn = inParam_optional_file ? castAndEscape(inParam_optional_file) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;
    const castedAs = as_optional_saveable_file_format ? castAndEscape(as_optional_saveable_file_format) : null;
    const valueForScriptAs = castedAs && typeof castedAs === 'string' && !castedAs.startsWith('{') && !castedAs.startsWith('date') ? `"${castedAs.replace(/"/g, "'")}"` : castedAs;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        save ${castedWindow}${inParam_optional_file ? ' in ' + valueForScriptIn : ''}${as_optional_saveable_file_format ? ' as ' + valueForScriptAs : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      window: target_window_required_string,
      in: inParam_optional_file || null,
      as: as_optional_saveable_file_format || null
    };
  }

  async printDocument(target_document_required_string, with_properties_optional_print_settings, print_dialog_optional_boolean) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const castedDocument = castAndEscape(target_document_required_string);
    const castedWith_properties = with_properties_optional_print_settings ? castAndEscape(with_properties_optional_print_settings) : null;
    const valueForScriptWith_properties = castedWith_properties && typeof castedWith_properties === 'string' && !castedWith_properties.startsWith('{') && !castedWith_properties.startsWith('date') ? `"${castedWith_properties.replace(/"/g, "'")}"` : castedWith_properties;
    const castedPrint_dialog = print_dialog_optional_boolean ? castAndEscape(print_dialog_optional_boolean) : null;
    const valueForScriptPrint_dialog = castedPrint_dialog && typeof castedPrint_dialog === 'string' && !castedPrint_dialog.startsWith('{') && !castedPrint_dialog.startsWith('date') ? `"${castedPrint_dialog.replace(/"/g, "'")}"` : castedPrint_dialog;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        print ${castedDocument}${with_properties_optional_print_settings ? ' with properties ' + valueForScriptWith_properties : ''}${print_dialog_optional_boolean ? ' print dialog ' + valueForScriptPrint_dialog : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      document: target_document_required_string,
      with_properties: with_properties_optional_print_settings || null,
      print_dialog: print_dialog_optional_boolean || null
    };
  }

  async printWindow(target_window_required_string, with_properties_optional_print_settings, print_dialog_optional_boolean) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedWindow = castAndEscape(target_window_required_string);
    const castedWith_properties = with_properties_optional_print_settings ? castAndEscape(with_properties_optional_print_settings) : null;
    const valueForScriptWith_properties = castedWith_properties && typeof castedWith_properties === 'string' && !castedWith_properties.startsWith('{') && !castedWith_properties.startsWith('date') ? `"${castedWith_properties.replace(/"/g, "'")}"` : castedWith_properties;
    const castedPrint_dialog = print_dialog_optional_boolean ? castAndEscape(print_dialog_optional_boolean) : null;
    const valueForScriptPrint_dialog = castedPrint_dialog && typeof castedPrint_dialog === 'string' && !castedPrint_dialog.startsWith('{') && !castedPrint_dialog.startsWith('date') ? `"${castedPrint_dialog.replace(/"/g, "'")}"` : castedPrint_dialog;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        print ${castedWindow}${with_properties_optional_print_settings ? ' with properties ' + valueForScriptWith_properties : ''}${print_dialog_optional_boolean ? ' print dialog ' + valueForScriptPrint_dialog : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      window: target_window_required_string,
      with_properties: with_properties_optional_print_settings || null,
      print_dialog: print_dialog_optional_boolean || null
    };
  }

  async quit(saving_optional_save_options) {
    const castedSaving = saving_optional_save_options ? castAndEscape(saving_optional_save_options) : null;

    const script = `
      tell application "Safari"
        quit${saving_optional_save_options ? ' saving ' + castedSaving : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      saving: saving_optional_save_options || null
    };
  }

  async countDocument() {


    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        count each document 
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async countTabOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedWindow = castAndEscape(target_window_required_string);

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        count each tab of ${castedWindow} 
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      window: target_window_required_string
    };
  }

  async countWindow() {


    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        count each window 
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async delete(direct_parameter_required_specifier) {
    if (direct_parameter_required_specifier === undefined || direct_parameter_required_specifier === null) {
      throw new Error("direct_parameter_required_specifier is required");
    }

    const castedDirect_parameter = direct_parameter_required_specifier ? castAndEscape(direct_parameter_required_specifier) : null;

    const script = `
      tell application "Safari"
        delete ${castedDirect_parameter}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_specifier || null
    };
  }

  async duplicate(direct_parameter_required_specifier, to_optional_location_specifier, with_properties_optional_record) {
    if (direct_parameter_required_specifier === undefined || direct_parameter_required_specifier === null) {
      throw new Error("direct_parameter_required_specifier is required");
    }

    const castedDirect_parameter = direct_parameter_required_specifier ? castAndEscape(direct_parameter_required_specifier) : null;
    const castedTo = to_optional_location_specifier ? castAndEscape(to_optional_location_specifier) : null;
    const castedWith_properties = with_properties_optional_record ? castAndEscape(with_properties_optional_record) : null;

    const script = `
      tell application "Safari"
        duplicate ${castedDirect_parameter}${to_optional_location_specifier ? ' to ' + castedTo : ''}${with_properties_optional_record ? ' with properties ' + castedWith_properties : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_specifier || null,
      to: to_optional_location_specifier || null,
      with_properties: with_properties_optional_record || null
    };
  }

  async exists(direct_parameter_required_any) {
    if (direct_parameter_required_any === undefined || direct_parameter_required_any === null) {
      throw new Error("direct_parameter_required_any is required");
    }

    const castedDirect_parameter = direct_parameter_required_any ? castAndEscape(direct_parameter_required_any) : null;

    const script = `
      tell application "Safari"
        exists ${castedDirect_parameter}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_any || null
    };
  }

  async makeDocument(at_optional_location_specifier, with_data_optional_any, with_properties_optional_text_url) {

    const castedAt = at_optional_location_specifier ? castAndEscape(at_optional_location_specifier) : null;
    const valueForScriptAt = castedAt && typeof castedAt === 'string' && !castedAt.startsWith('{') && !castedAt.startsWith('date') ? `"${castedAt.replace(/"/g, "'")}"` : castedAt;
    const castedWith_data = with_data_optional_any ? castAndEscape(with_data_optional_any) : null;
    const valueForScriptWith_data = castedWith_data && typeof castedWith_data === 'string' && !castedWith_data.startsWith('{') && !castedWith_data.startsWith('date') ? `"${castedWith_data.replace(/"/g, "'")}"` : castedWith_data;
    const castedWith_properties_optional_text_url = with_properties_optional_text_url ? castAndEscape(with_properties_optional_text_url) : null;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        make new document ${at_optional_location_specifier ? ' at ' + valueForScriptAt : ''}${with_data_optional_any ? ' with data ' + valueForScriptWith_data : ''}${buildPropertiesRecord([{param: 'with_properties_optional_text_url', prop: 'URL', value: with_properties_optional_text_url, type: 'text'}])}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      at: at_optional_location_specifier || null,
      with_data: with_data_optional_any || null,
      url: with_properties_optional_text_url || null
    };
  }

  async makeTabOfWindow(at_required_location_specifier_window, with_data_optional_any, with_properties_optional_text_url) {
    if (!at_required_location_specifier_window || typeof at_required_location_specifier_window !== "string") {
      throw new Error("at_required_location_specifier_window is required and must be a string");
    }

    const castedWindow = castAndEscape(at_required_location_specifier_window);
    const castedWith_data = with_data_optional_any ? castAndEscape(with_data_optional_any) : null;
    const valueForScriptWith_data = castedWith_data && typeof castedWith_data === 'string' && !castedWith_data.startsWith('{') && !castedWith_data.startsWith('date') ? `"${castedWith_data.replace(/"/g, "'")}"` : castedWith_data;
    const castedWith_properties_optional_text_url = with_properties_optional_text_url ? castAndEscape(with_properties_optional_text_url) : null;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        make new tab at ${castedWindow} ${with_data_optional_any ? ' with data ' + valueForScriptWith_data : ''}${buildPropertiesRecord([{param: 'with_properties_optional_text_url', prop: 'URL', value: with_properties_optional_text_url, type: 'text'}])}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      window: at_required_location_specifier_window,
      with_data: with_data_optional_any || null,
      url: with_properties_optional_text_url || null
    };
  }

  async makeWindow(at_optional_location_specifier, with_data_optional_any, with_properties_optional_integer_index, with_properties_optional_tab_current_tab, with_properties_optional_boolean_zoomed, with_properties_optional_boolean_miniaturized, with_properties_optional_boolean_visible, with_properties_optional_rectangle_bounds) {

    const castedAt = at_optional_location_specifier ? castAndEscape(at_optional_location_specifier) : null;
    const valueForScriptAt = castedAt && typeof castedAt === 'string' && !castedAt.startsWith('{') && !castedAt.startsWith('date') ? `"${castedAt.replace(/"/g, "'")}"` : castedAt;
    const castedWith_data = with_data_optional_any ? castAndEscape(with_data_optional_any) : null;
    const valueForScriptWith_data = castedWith_data && typeof castedWith_data === 'string' && !castedWith_data.startsWith('{') && !castedWith_data.startsWith('date') ? `"${castedWith_data.replace(/"/g, "'")}"` : castedWith_data;
    const castedWith_properties_optional_integer_index = with_properties_optional_integer_index ? castAndEscape(with_properties_optional_integer_index) : null;
    const castedWith_properties_optional_tab_current_tab = with_properties_optional_tab_current_tab ? castAndEscape(with_properties_optional_tab_current_tab) : null;
    const castedWith_properties_optional_boolean_zoomed = with_properties_optional_boolean_zoomed ? castAndEscape(with_properties_optional_boolean_zoomed) : null;
    const castedWith_properties_optional_boolean_miniaturized = with_properties_optional_boolean_miniaturized ? castAndEscape(with_properties_optional_boolean_miniaturized) : null;
    const castedWith_properties_optional_boolean_visible = with_properties_optional_boolean_visible ? castAndEscape(with_properties_optional_boolean_visible) : null;
    const castedWith_properties_optional_rectangle_bounds = with_properties_optional_rectangle_bounds ? castAndEscape(with_properties_optional_rectangle_bounds) : null;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        make new window ${at_optional_location_specifier ? ' at ' + valueForScriptAt : ''}${with_data_optional_any ? ' with data ' + valueForScriptWith_data : ''}${buildPropertiesRecord([{param: 'with_properties_optional_integer_index', prop: 'index', value: with_properties_optional_integer_index, type: 'integer'}, {param: 'with_properties_optional_tab_current_tab', prop: 'current tab', value: with_properties_optional_tab_current_tab, type: 'tab'}, {param: 'with_properties_optional_boolean_zoomed', prop: 'zoomed', value: with_properties_optional_boolean_zoomed, type: 'boolean'}, {param: 'with_properties_optional_boolean_miniaturized', prop: 'miniaturized', value: with_properties_optional_boolean_miniaturized, type: 'boolean'}, {param: 'with_properties_optional_boolean_visible', prop: 'visible', value: with_properties_optional_boolean_visible, type: 'boolean'}, {param: 'with_properties_optional_rectangle_bounds', prop: 'bounds', value: with_properties_optional_rectangle_bounds, type: 'rectangle'}])}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      at: at_optional_location_specifier || null,
      with_data: with_data_optional_any || null,
      index: with_properties_optional_integer_index || null,
      current_tab: with_properties_optional_tab_current_tab || null,
      zoomed: with_properties_optional_boolean_zoomed || null,
      miniaturized: with_properties_optional_boolean_miniaturized || null,
      visible: with_properties_optional_boolean_visible || null,
      bounds: with_properties_optional_rectangle_bounds || null
    };
  }

  async move(direct_parameter_required_specifier, to_required_location_specifier) {
    if (direct_parameter_required_specifier === undefined || direct_parameter_required_specifier === null) {
      throw new Error("direct_parameter_required_specifier is required");
    }

    if (to_required_location_specifier === undefined || to_required_location_specifier === null) {
      throw new Error("to_required_location_specifier is required");
    }

    const castedDirect_parameter = direct_parameter_required_specifier ? castAndEscape(direct_parameter_required_specifier) : null;
    const castedTo = to_required_location_specifier ? castAndEscape(to_required_location_specifier) : null;

    const script = `
      tell application "Safari"
        move ${castedDirect_parameter} to ${castedTo}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_specifier || null,
      to: to_required_location_specifier || null
    };
  }

  async getNameOfApplication() {
    const script = `
      tell application "Safari"
        return name
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script
    };
  }

  async getFrontmostOfApplication() {
    const script = `
      tell application "Safari"
        return frontmost
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script
    };
  }

  async getVersionOfApplication() {
    const script = `
      tell application "Safari"
        return version
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script
    };
  }

  async getNameOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return name of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async getModifiedOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return modified of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async getFileOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return file of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async getNameOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return name of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getIdOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return id of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getIndexOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return index of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setIndexOfWindow(target_window_required_string, value_required_integer) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_integer === undefined || value_required_integer === null) {
      throw new Error("value_required_integer is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_integer, 'integer');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set index of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_integer,
      script: script,
      window: target_window_required_string
    };
  }

  async getBoundsOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return bounds of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setBoundsOfWindow(target_window_required_string, value_required_rectangle) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_rectangle === undefined || value_required_rectangle === null) {
      throw new Error("value_required_rectangle is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_rectangle, 'rectangle');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set bounds of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_rectangle,
      script: script,
      window: target_window_required_string
    };
  }

  async getCloseableOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return closeable of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getMiniaturizableOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return miniaturizable of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getMiniaturizedOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return miniaturized of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setMiniaturizedOfWindow(target_window_required_string, value_required_boolean) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_boolean === undefined || value_required_boolean === null) {
      throw new Error("value_required_boolean is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_boolean, 'boolean');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set miniaturized of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_boolean,
      script: script,
      window: target_window_required_string
    };
  }

  async getResizableOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return resizable of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getVisibleOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return visible of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setVisibleOfWindow(target_window_required_string, value_required_boolean) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_boolean === undefined || value_required_boolean === null) {
      throw new Error("value_required_boolean is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_boolean, 'boolean');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set visible of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_boolean,
      script: script,
      window: target_window_required_string
    };
  }

  async getZoomableOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return zoomable of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async getZoomedOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return zoomed of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setZoomedOfWindow(target_window_required_string, value_required_boolean) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_boolean === undefined || value_required_boolean === null) {
      throw new Error("value_required_boolean is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_boolean, 'boolean');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set zoomed of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_boolean,
      script: script,
      window: target_window_required_string
    };
  }

  async getDocumentOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return document of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async addReadingListItem(direct_parameter_required_text, and_preview_text_optional_text, with_title_optional_text) {
    if (direct_parameter_required_text === undefined || direct_parameter_required_text === null) {
      throw new Error("direct_parameter_required_text is required");
    }

    const castedDirect_parameter = direct_parameter_required_text ? castAndEscape(direct_parameter_required_text) : null;
    const castedAnd_preview_text = and_preview_text_optional_text ? castAndEscape(and_preview_text_optional_text) : null;
    const castedWith_title = with_title_optional_text ? castAndEscape(with_title_optional_text) : null;

    const script = `
      tell application "Safari"
        add reading list item ${castedDirect_parameter}${and_preview_text_optional_text ? ' and preview text ' + castedAnd_preview_text : ''}${with_title_optional_text ? ' with title ' + castedWith_title : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_text || null,
      and_preview_text: and_preview_text_optional_text || null,
      with_title: with_title_optional_text || null
    };
  }

  async doJavascriptDocument(target_document_required_string, inParam_optional_document) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const castedDocument = castAndEscape(target_document_required_string);
    const castedIn = inParam_optional_document ? castAndEscape(inParam_optional_document) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        do JavaScript ${castedDocument}${inParam_optional_document ? ' in ' + valueForScriptIn : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      document: target_document_required_string,
      in: inParam_optional_document || null
    };
  }

  async doJavascriptTabOfWindow(target_tab_required_string, target_window_required_string, inParam_optional_document) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const castedTab = castAndEscape(target_tab_required_string);
    const castedWindow = castAndEscape(target_window_required_string);
    const castedIn = inParam_optional_document ? castAndEscape(inParam_optional_document) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        do JavaScript ${castedTab} of ${castedWindow}${inParam_optional_document ? ' in ' + valueForScriptIn : ''}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string,
      in: inParam_optional_document || null
    };
  }


  async searchTheWebDocument(target_document_required_string, inParam_optional_document, forParam_required_text) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }
    if (forParam_required_text === undefined || forParam_required_text === null) {
      throw new Error("forParam_required_text is required");
    }

    const castedDocument = castAndEscape(target_document_required_string);
    const castedIn = inParam_optional_document ? castAndEscape(inParam_optional_document) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;
    const castedFor = forParam_required_text ? castAndEscape(forParam_required_text) : null;
    const valueForScriptFor = castedFor && typeof castedFor === 'string' && !castedFor.startsWith('{') && !castedFor.startsWith('date') ? `"${castedFor.replace(/"/g, "'")}"` : castedFor;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        search the web ${castedDocument}${inParam_optional_document ? ' in ' + valueForScriptIn : ''} for ${valueForScriptFor}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      document: target_document_required_string,
      in: inParam_optional_document || null,
      for: forParam_required_text || null
    };
  }

  async searchTheWebTabOfWindow(target_tab_required_string, target_window_required_string, inParam_optional_document, forParam_required_text) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (forParam_required_text === undefined || forParam_required_text === null) {
      throw new Error("forParam_required_text is required");
    }

    const castedTab = castAndEscape(target_tab_required_string);
    const castedWindow = castAndEscape(target_window_required_string);
    const castedIn = inParam_optional_document ? castAndEscape(inParam_optional_document) : null;
    const valueForScriptIn = castedIn && typeof castedIn === 'string' && !castedIn.startsWith('{') && !castedIn.startsWith('date') ? `"${castedIn.replace(/"/g, "'")}"` : castedIn;
    const castedFor = forParam_required_text ? castAndEscape(forParam_required_text) : null;
    const valueForScriptFor = castedFor && typeof castedFor === 'string' && !castedFor.startsWith('{') && !castedFor.startsWith('date') ? `"${castedFor.replace(/"/g, "'")}"` : castedFor;

    // Helper function to build properties record from individual property parameters
    function buildPropertiesRecord(propertyParams) {
      const definedProps = propertyParams.filter(p => p.value !== undefined && p.value !== null && p.value !== '');
      if (definedProps.length === 0) return '';
      const propStrings = definedProps.map(p => {
        const castedValue = castAndEscape(p.value, p.type || null);
        // For strings that got escaped, wrap in quotes and replace inner quotes
        if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
          return `${p.prop}:"${castedValue.replace(/"/g, "'")}"`;
        }
        // For numbers, booleans, lists, records, dates - no quotes
        return `${p.prop}:${castedValue}`;
      });
      return ` with properties {${propStrings.join(', ')}}`;
    }

    const script = `
      tell application "Safari"
        search the web ${castedTab} of ${castedWindow}${inParam_optional_document ? ' in ' + valueForScriptIn : ''} for ${valueForScriptFor}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string,
      in: inParam_optional_document || null,
      for: forParam_required_text || null
    };
  }

  async showBookmarks() {
    const script = `
      tell application "Safari"
        show bookmarks
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async showExtensionsPreferences(direct_parameter_required_text) {
    if (direct_parameter_required_text === undefined || direct_parameter_required_text === null) {
      throw new Error("direct_parameter_required_text is required");
    }

    const castedDirect_parameter = direct_parameter_required_text ? castAndEscape(direct_parameter_required_text) : null;

    const script = `
      tell application "Safari"
        show extensions preferences ${castedDirect_parameter}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_text || null
    };
  }

  async dispatchMessageToExtension(direct_parameter_required_any) {
    if (direct_parameter_required_any === undefined || direct_parameter_required_any === null) {
      throw new Error("direct_parameter_required_any is required");
    }

    const castedDirect_parameter = direct_parameter_required_any ? castAndEscape(direct_parameter_required_any) : null;

    const script = `
      tell application "Safari"
        dispatch message to extension ${castedDirect_parameter}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script,
      direct_parameter: direct_parameter_required_any || null
    };
  }

  async syncAllPlistToDisk() {
    const script = `
      tell application "Safari"
        sync all plist to disk
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async showPrivacyReport() {
    const script = `
      tell application "Safari"
        show privacy report
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async showCreditCardSettings() {
    const script = `
      tell application "Safari"
        show credit card settings
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: result,
      script: script
    };
  }

  async getCurrentTabOfWindow(target_window_required_string) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return current tab of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      window: target_window_required_string
    };
  }

  async setCurrentTabOfWindow(target_window_required_string, value_required_tab) {
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_tab === undefined || value_required_tab === null) {
      throw new Error("value_required_tab is required");
    }

    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_tab, 'tab');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set current tab of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_tab,
      script: script,
      window: target_window_required_string
    };
  }

  async getSourceOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return source of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async getUrlOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return URL of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async setUrlOfDocument(target_document_required_string, value_required_text) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }
    if (value_required_text === undefined || value_required_text === null) {
      throw new Error("value_required_text is required");
    }

    const castedDocument = castAndEscape(target_document_required_string, 'string');
    const castedValue = castAndEscape(value_required_text, 'text');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set URL of ${castedDocument} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_text,
      script: script,
      document: target_document_required_string
    };
  }

  async getTextOfDocument(target_document_required_string) {
    if (!target_document_required_string || typeof target_document_required_string !== "string") {
      throw new Error("target_document_required_string is required and must be a string");
    }

    const escapedDocument = escapeForAppleScript(target_document_required_string);

    const script = `
      tell application "Safari"
        return text of ${escapedDocument}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      document: target_document_required_string
    };
  }

  async getSourceOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return source of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getUrlOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return URL of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async setUrlOfTabOfWindow(target_tab_required_string, target_window_required_string, value_required_text) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }
    if (value_required_text === undefined || value_required_text === null) {
      throw new Error("value_required_text is required");
    }

    const castedTab = castAndEscape(target_tab_required_string, 'string');
    const castedWindow = castAndEscape(target_window_required_string, 'string');
    const castedValue = castAndEscape(value_required_text, 'text');
    // Determine value format for AppleScript
    let valueForScript;
    if (typeof castedValue === 'string' && !castedValue.startsWith('{') && !castedValue.startsWith('date')) {
      valueForScript = `"${castedValue}"`; // Wrap strings in quotes
    } else {
      valueForScript = castedValue; // Use as-is for numbers, booleans, lists, records
    }

    const script = `
      tell application "Safari"
        set URL of ${castedTab} of ${castedWindow} to ${valueForScript}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      message: "Property set successfully",
      value: value_required_text,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getIndexOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return index of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getTextOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return text of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getVisibleOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return visible of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getNameOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return name of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }

  async getPidOfTabOfWindow(target_tab_required_string, target_window_required_string) {
    if (!target_tab_required_string || typeof target_tab_required_string !== "string") {
      throw new Error("target_tab_required_string is required and must be a string");
    }
    if (!target_window_required_string || typeof target_window_required_string !== "string") {
      throw new Error("target_window_required_string is required and must be a string");
    }

    const escapedTab = escapeForAppleScript(target_tab_required_string);
    const escapedWindow = escapeForAppleScript(target_window_required_string);

    const script = `
      tell application "Safari"
        return pid of ${escapedTab} of ${escapedWindow}
      end tell
    `;

    const result = await executeAppleScript(script);
    return {
      success: result !== "Error",
      value: result,
      script: script,
      tab: target_tab_required_string,
      window: target_window_required_string
    };
  }





  sendResponse(response) {
    const responseStr = JSON.stringify(response);
    console.error("Sending response:", response.method || 'result', response.id);
    process.stdout.write(responseStr + '\n');
  }
}

// Start the server
async function startServer() {
  console.error("Testing Safari availability...");
  await checkSafariAvailable();
  
  console.error("Creating Safari MCP server...");
  const server = new SafariMCPServer();
  
  console.error("Safari AppleScript MCP server running on stdio");
  
  // Keep the process alive
  process.on('SIGINT', () => {
    console.error("Shutting down Safari AppleScript MCP server");
    process.exit(0);
  });
  
  process.on('SIGTERM', () => {
    console.error("Shutting down Safari AppleScript MCP server");
    process.exit(0);
  });
}

startServer().catch(error => {
  console.error("Fatal error starting server:", error);
  process.exit(1);
});
