/**
 * Error handling utilities following Node.js best practices
 */

/**
 * Custom application error class
 */
class AppError extends Error {
  constructor(name, message, statusCode = 500, isOperational = true) {
    super(message);
    
    this.name = name;
    this.statusCode = statusCode;
    this.isOperational = isOperational;
    
    // Maintain stack trace
    Error.captureStackTrace(this, AppError);
  }
}

/**
 * Common error types
 */
const ErrorTypes = {
  AUTHENTICATION_ERROR: 'AuthenticationError',
  AUTHORIZATION_ERROR: 'AuthorizationError',
  VALIDATION_ERROR: 'ValidationError',
  NOT_FOUND_ERROR: 'NotFoundError',
  GRAPH_API_ERROR: 'GraphApiError',
  CONFIGURATION_ERROR: 'ConfigurationError'
};

/**
 * Create a standardized error response for MCP tools
 * @param {Error} error - The error object
 * @returns {Object} - MCP error response
 */
function createMcpErrorResponse(error) {
  const isOperational = error.isOperational || false;
  
  return {
    content: [{
      type: 'text',
      text: isOperational ? error.message : 'An unexpected error occurred'
    }],
    isError: true
  };
}

/**
 * Create authentication error
 * @param {string} message - Error message
 * @returns {AppError} - Authentication error
 */
function createAuthError(message = 'Authentication required') {
  return new AppError(ErrorTypes.AUTHENTICATION_ERROR, message, 401);
}

/**
 * Create validation error
 * @param {string} message - Error message  
 * @returns {AppError} - Validation error
 */
function createValidationError(message = 'Invalid input provided') {
  return new AppError(ErrorTypes.VALIDATION_ERROR, message, 400);
}

/**
 * Create Graph API error
 * @param {string} message - Error message
 * @param {number} statusCode - HTTP status code
 * @returns {AppError} - Graph API error
 */
function createGraphApiError(message, statusCode = 500) {
  return new AppError(ErrorTypes.GRAPH_API_ERROR, message, statusCode);
}

module.exports = {
  AppError,
  ErrorTypes,
  createMcpErrorResponse,
  createAuthError,
  createValidationError,
  createGraphApiError
};