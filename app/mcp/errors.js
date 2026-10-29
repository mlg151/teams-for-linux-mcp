'use strict';

class McpOperationError extends Error {
  constructor(code, message, details = {}) {
    super(message);
    this.code = code;
    this.details = details;
  }
}

module.exports = { McpOperationError };
