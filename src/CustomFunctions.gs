/**
 * Custom Functions for Google Sheets
 * These functions can be used directly in cells like =SENTIMENTSCORE("text")
 * Perfect for poker earnings tracking and currency conversions!
 */

/**
 * Converts amount from one currency to another using live rates
 * 
 * @param {number} amount The amount to convert
 * @param {string} fromCurrency Source currency code (e.g., "USD")
 * @param {string} toCurrency Target currency code (e.g., "EUR")
 * @return {number} Converted amount
 * @customfunction
 */
function CONVERTCURRENCY(amount, fromCurrency, toCurrency) {
  if (typeof amount !== 'number' || !fromCurrency || !toCurrency) return '#ERROR';
  
  try {
    const url = `https://api.exchangerate-api.com/v4/latest/${fromCurrency.toUpperCase()}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    const rate = data.rates[toCurrency.toUpperCase()];
    if (!rate) return '#INVALID_CURRENCY';
    
    return amount * rate;
  } catch (error) {
    return '#API_ERROR';
  }
}

/**
 * Gets the current exchange rate between two currencies
 * 
 * @param {string} fromCurrency Source currency code
 * @param {string} toCurrency Target currency code
 * @return {number} Exchange rate
 * @customfunction
 */
function EXCHANGERATE(fromCurrency, toCurrency) {
  if (!fromCurrency || !toCurrency) return '#ERROR';
  
  try {
    const url = `https://api.exchangerate-api.com/v4/latest/${fromCurrency.toUpperCase()}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    const rate = data.rates[toCurrency.toUpperCase()];
    return rate || '#INVALID_CURRENCY';
  } catch (error) {
    return '#API_ERROR';
  }
}

/**
 * Calculates ROI (Return on Investment) percentage
 * 
 * @param {number} profit The profit amount
 * @param {number} investment The initial investment
 * @return {string} ROI percentage with sign
 * @customfunction
 */
function CALCULATEroi(profit, investment) {
  if (typeof profit !== 'number' || typeof investment !== 'number') return '#ERROR';
  if (investment === 0) return 'N/A';
  
  const roi = (profit / investment) * 100;
  const sign = roi >= 0 ? '+' : '';
  
  return sign + roi.toFixed(2) + '%';
}

/**
 * Calculates a simple sentiment score for text (-1 to 1)
 * 
 * @param {string} text The text to analyze
 * @return {number} Sentiment score between -1 (negative) and 1 (positive)
 * @customfunction
 */
function SENTIMENTSCORE(text) {
  if (!text || typeof text !== 'string') return 0;
  
  const positive = ['good', 'great', 'excellent', 'amazing', 'wonderful', 'love', 'best', 'awesome', 'fantastic', 'perfect'];
  const negative = ['bad', 'terrible', 'awful', 'hate', 'worst', 'poor', 'horrible', 'disappointing', 'useless'];
  
  const lowerText = text.toLowerCase();
  let score = 0;
  
  positive.forEach(word => {
    const regex = new RegExp('\\b' + word + '\\b', 'gi');
    const matches = lowerText.match(regex);
    if (matches) score += matches.length;
  });
  
  negative.forEach(word => {
    const regex = new RegExp('\\b' + word + '\\b', 'gi');
    const matches = lowerText.match(regex);
    if (matches) score -= matches.length;
  });
  
  return Math.max(-1, Math.min(1, score * 0.2));
}

/**
 * Extracts email addresses from text
 * 
 * @param {string} text The text to search for emails
 * @return {string} Comma-separated list of emails found
 * @customfunction
 */
function EXTRACTEMAILS(text) {
  if (!text || typeof text !== 'string') return '';
  
  const emailRegex = /[\w.-]+@[\w.-]+\.\w+/g;
  const emails = text.match(emailRegex);
  
  return emails ? emails.join(', ') : '';
}

/**
 * Calculates reading time for text (average reading speed: 200 words/min)
 * 
 * @param {string} text The text to analyze
 * @return {string} Estimated reading time
 * @customfunction
 */
function READINGTIME(text) {
  if (!text || typeof text !== 'string') return '0 min';
  
  const words = text.trim().split(/\s+/).length;
  const minutes = Math.ceil(words / 200);
  
  return minutes + ' min';
}

/**
 * Generates a random string ID
 * 
 * @param {number} length Length of the ID (default: 8)
 * @return {string} Random alphanumeric ID
 * @customfunction
 */
function RANDOMID(length) {
  const len = length || 8;
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  
  for (let i = 0; i < len; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  
  return result;
}

/**
 * Calculates percentage change between two numbers
 * 
 * @param {number} oldValue The original value
 * @param {number} newValue The new value
 * @return {string} Percentage change with + or - sign
 * @customfunction
 */
function PERCENTCHANGE(oldValue, newValue) {
  if (typeof oldValue !== 'number' || typeof newValue !== 'number') return '#ERROR';
  if (oldValue === 0) return 'N/A';
  
  const change = ((newValue - oldValue) / oldValue) * 100;
  const sign = change >= 0 ? '+' : '';
  
  return sign + change.toFixed(2) + '%';
}

/**
 * Converts text to title case
 * 
 * @param {string} text The text to convert
 * @return {string} Title cased text
 * @customfunction
 */
function TITLECASE(text) {
  if (!text || typeof text !== 'string') return '';
  
  return text.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

/**
 * Removes all non-numeric characters from text
 * 
 * @param {string} text The text to clean
 * @return {number} Numeric value extracted
 * @customfunction
 */
function EXTRACTNUMBER(text) {
  if (!text) return 0;
  
  const cleaned = String(text).replace(/[^\d.-]/g, '');
  const number = parseFloat(cleaned);
  
  return isNaN(number) ? 0 : number;
}

/**
 * Calculates the nth Fibonacci number
 * 
 * @param {number} n The position in the Fibonacci sequence
 * @return {number} The nth Fibonacci number
 * @customfunction
 */
function FIBONACCI(n) {
  if (typeof n !== 'number' || n < 0) return '#ERROR';
  if (n <= 1) return n;
  
  let a = 0, b = 1;
  for (let i = 2; i <= n; i++) {
    const temp = a + b;
    a = b;
    b = temp;
  }
  
  return b;
}

/**
 * Checks if a string is a valid email address
 * 
 * @param {string} email The email address to validate
 * @return {boolean} TRUE if valid email, FALSE otherwise
 * @customfunction
 */
function ISVALIDEMAIL(email) {
  if (!email || typeof email !== 'string') return false;
  
  const emailRegex = /^[\w.-]+@[\w.-]+\.\w+$/;
  return emailRegex.test(email.trim());
}

/**
 * Counts words in text
 * 
 * @param {string} text The text to count words in
 * @return {number} Number of words
 * @customfunction
 */
function WORDCOUNT(text) {
  if (!text || typeof text !== 'string') return 0;
  
  return text.trim().split(/\s+/).filter(word => word.length > 0).length;
}
