/**
 * A function to be executed.
 * @callback caller
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 */

/**
 * A callback to be executed on each iteration with the value retrieved from the iterator as the argument.
 * @callback iteratorCallback
 * @param {*} value - The value retrieved from the iterator.
 * @throws {StopError} - If the error is an instance of StopError, the execution of the timeout function will be stopped.
 */

/**
 * An iterator object.
 * @typedef {Object} Iterators
 * @property {Array} array - The array to be iterated over.
 * @property {GoogleAppsScript.Spreadsheet.Range} range - The range object to iterate over.
 * @property {requestCallback} request - A callback to request that paginated results be returned.
 */

/**
 * A callback to request that paginated results be returned.
 * @callback requestCallback
 * @param {string} pageToken - The page token to be used to retrieve the next page of results.
 * @throws {StopError} - If the error is an instance of StopError, the execution of the timeout function will be stopped.
 */

/**
 * An object containing options for the timeout function.
 * @typedef {Object} Options
 * @property {number} timeout - The maximum amount of time (in milliseconds) that the function can run before it is interrupted.
 * @property {number} delay - The amount of time (in milliseconds) to wait before restarting the function.
 * @property {number} start - The start time (in milliseconds) of the function.
 * @property {boolean} debug - Whether or not to enable debug logging.
 */

/**
 * The timeout() function executes the specified code on a set of values sourced from an iterable object. Sets a timer and restarts the specified code when the execution time approaches the Google Apps Script time limit.
 * @param {caller} caller - The function to be executed.
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 * @param {Iterators} iterator - A union containing the values ​​to be processed.
 * @param {Array} [iterator.array] - The array to be iterated over.
 * @param {GoogleAppsScript.Spreadsheet.Range} [iterator.range] - The range object to iterate over.
 * @param {requestCallback} [iterator.request] - A callback to request that paginated results be returned.
 * @param {iteratorCallback} callback - The function to be executed for each value in the iterable object.
 * @param {Options} [options] - The options object.
 * @param {number} [options.timeout=300000] - The maximum amount of time (in milliseconds) that the function can run before it is interrupted.
 * @param {number} [options.delay=60000] - The amount of time (in milliseconds) to wait before restarting the function.
 * @param {number} [options.start=Date.now()] - The start time (in milliseconds) of the function.
 * @param {boolean} [options.debug=false] - Whether or not to enable debug logging.
 * @returns {GoogleAppsScript.Script.Trigger | null} - a GoogleAppsScript.Script.Trigger or null
 */
function timeout(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    // Default timeout of 5 minutes (in milliseconds)
    timeout: 300000, // 5 * 60 * 1000
    // Default delay of 1 minute (in milliseconds)
    delay: 60000, // 1 * 60 * 1000
    // Record the start time
    start: Date.now(),
    // Default debug logging to off
    debug: false
  }, options);

  if (typeof caller !== 'function') {
    throw new TypeError('caller is not function.');
  }
  if (!caller.name) {
    throw new ReferenceError('caller.name is not defined.');
  }
  if (typeof callback !== 'function') {
    throw new TypeError('callback is not function.');
  }
  if (!iterator) {
    throw new ReferenceError('iterator is not defined.');
  }
  if ([iterator.array, iterator.range, iterator.request].filter((e) => !!e).length > 1) {
    throw new TypeError('iterator cannot be both array and range and request.');
  }
  if (iterator.array) {
    iterator = createRestartableIteratorFromArray_(iterator.array);
  } else if (iterator.range) {
    iterator = createRestartableIteratorFromRange_(iterator.range);
  } else if (iterator.request) {
    iterator = createRestartableIteratorFromPageTokenResponse_(iterator.request);
  }

  if (event && event.triggerUid) {
    const value = PropertiesService.getUserProperties()
      .getProperty(`${ScriptApp.getScriptId()}/${event.triggerUid}`);

    if (value === null) {
      throw new Error('Cannot find property for %s.', event.triggerUid);
    }

    const property = JSON.parse(value);
    console.info('Restore property.', property);

    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === event.triggerUid) {
        ScriptApp.deleteTrigger(trigger);
        console.log('Delete trigger: %s', event.triggerUid);

        PropertiesService.getUserProperties()
          .deleteProperty(`${ScriptApp.getScriptId()}/${event.triggerUid}`);
        console.log('Delete property.');
      }
    }

    Object.assign(iterator, property);
  }

  for (const value of iterator) {
    console.info(`iterator[${iterator.index}/${iterator.length}]`);

    try {
      console.time(caller.name);

      callback(value);

      console.timeEnd(caller.name);
    } catch (error) {
      if (error instanceof StopError) {
        console.info('StopError catched.');

        return null;
      }

      throw error;
    }

    const progress = Date.now() - options.start;
    if (progress > options.timeout) {
      break;
    }
  }

  if ((iterator !== undefined && iterator.length === undefined) || (iterator.index < iterator.length)) {
    const trigger = ScriptApp.newTrigger(caller.name)
      .timeBased()
      .after(options.delay)
      .create();

    console.info('Create trigger. Function name: %s Unique ID: %s', trigger.getHandlerFunction(), trigger.getUniqueId());

    const property = {
      index: iterator.index,
      length: iterator.length
    };

    PropertiesService.getUserProperties()
      .setProperty(`${ScriptApp.getScriptId()}/${trigger.getUniqueId()}`, JSON.stringify(property));
    console.info('Save property.', property);

    return trigger;
  }

  return null;
}

/**
 * Execute the specified function in parallel with timeout handling.
 * @param {caller} caller - The function to be executed.
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 * @param {Iterators} iterator - A union containing the values ​​to be processed.
 * @param {Array} [iterator.array] - The array to be iterated over.
 * @param {GoogleAppsScript.Spreadsheet.Range} [iterator.range] - The range object to iterate over.
 * @param {iteratorCallback} callback - The function to be executed for each value in the iterable object.
 * @param {Options} [options] - The options object.
 * @param {number} [options.split=4] - The number of times to split the iterable object.
 * @param {number} [options.timeout=300000] - The maximum amount of time (in milliseconds) that the function can run before it is interrupted.
 * @param {number} [options.delay=60000] - The amount of time (in milliseconds) to wait before restarting the function.
 * @param {number} [options.start=Date.now()] - The start time (in milliseconds) of the function.
 * @param {boolean} [options.debug=false] - Whether or not to enable debug logging.
 */
function timeoutInParallel(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    // Default split value 
    split: 4,
    // Default timeout of 5 minutes (in milliseconds)
    timeout: 300000, // 5 * 60 * 1000
    // Default delay of 1 minute (in milliseconds)
    delay: 60000, // 1 * 60 * 1000
    // Record the start time
    start: Date.now(),
    // Default debug logging to off
    debug: false
  }, options);

  if (!iterator) {
    throw new ReferenceError('iterator is not defined.');
  }
  // Check if the iterator is a request object.
  if (iterator.request) {
    throw new TypeError('iterators cannot be request.');
  }
  if ([iterator.array, iterator.range].filter((e) => !!e).length > 1) {
    throw new TypeError('iterator cannot be both array and range.');
  }
  if (iterator.array) {
    iterator = createRestartableIteratorFromArray_(iterator.array);
  } else if (iterator.range) {
    iterator = createRestartableIteratorFromRange_(iterator.range);
  }

  // Check if the split option is valid (cannot be greater than 4)
  if (options.split > 4) {
    throw new RangeError('options.split cannot be greater than 4.');
  }

  // If no event is provided, initiate the splitting process
  if (!event || !event.triggerUid) {
    // Check if splitting would exceed the maximum trigger limit
    if (Math.pow(2, options.split) > (20 - ScriptApp.getProjectTriggers().length)) {
      throw new Error('Limit will be exceede.');
    }

    // Create split triggers for parallel execution
    createSplitTrigger_(options.split, iterator.index, iterator.length, caller, options);
  }

  // If an event is provided, handle the timeout logic
  timeout(caller, event, iterator, callback, options);
}

/**
 * Recursively creates split triggers for time-based execution of a function. Designed to break down a larger task into smaller, scheduled executions.
 * @param {number} split - The number of splits remaining, used to determine recursion depth.
 * @param {number} index - The starting index for the current task segment.
 * @param {number} length - The total length of the task.
 * @param {caller} caller - The function to be executed by the triggers.
 * @param {Options} options - An object containing configuration options for the triggers:
 */
function createSplitTrigger_(split, index, length, caller, options) {
  // Calculate the length of the current section
  const section = length - index;

  // Calculate the midpoint of the current section
  const left = Math.floor(section / 2);

  // If further splits are possible, recursively create triggers for the two halves
  if (split > 0 && left > 1) {
    createSplitTrigger_(split - 1, index, index + left, caller, options);
    createSplitTrigger_(split - 1, index + left, length, caller, options);
  } else {
    // Base case: create a trigger for the current segment
    const trigger = ScriptApp.newTrigger(caller.name)
      .timeBased()
      .after(options.delay)
      .create();
    console.info('Create trigger. Function name: %s Unique ID: %s', trigger.getHandlerFunction(), trigger.getUniqueId());

    // Store the start and end indices for the segment as a property associated with the trigger
    const property = {
      index: index,
      length: length
    };

    // Store the properties associated with the trigger
    PropertiesService.getUserProperties()
      .setProperty(`${ScriptApp.getScriptId()}/${trigger.getUniqueId()}`, JSON.stringify(property));
    console.info('Save property.', property);
  }
}

/**
 * Creates a restartable iterator from a provided range object. The iterator allows sequential access to values within the range, and offers the ability to 'restart' iteration from the beginning.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object to iterate over.
 * @returns {Object} A sealed iterator object.
 */
function createRestartableIteratorFromRange_(range) {
  const iterator = new Object();

  Object.defineProperties(iterator, {
    index: {
      configurable: false,
      enumerable: true,
      value: 0,
      writable: true
    },
    length: {
      configurable: false,
      enumerable: true,
      value: range.getNumRows(),
      writable: true
    },
    [Symbol.iterator]: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        return this;
      }
    },
    next: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        if (this.index < this.length) {
          return {
            value: range.offset(this.index++, 0, 1),
            done: false
          };
        }
        return {
          done: true
        };
      }
    },
    return: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (value) {
        return {
          value: value,
          done: true
        };
      }
    },
    throw: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (exception) {
        return {
          done: true
        };
      }
    }
  });

  return Object.seal(iterator);
}

/**
 * Creates a restartable iterator from a provided 2D array. The iterator allows sequential access to the inner arrays within the main array, and offers the ability to 'restart' iteration from the beginning.
 * @param {Object[][]} array - The 2D array to iterate over.
 * @returns {Object} A sealed iterator object.
 */
function createRestartableIteratorFromArray_(array) {
  const iterator = new Object();

  Object.defineProperties(iterator, {
    index: {
      configurable: false,
      enumerable: true,
      value: 0,
      writable: true
    },
    length: {
      configurable: false,
      enumerable: true,
      value: array.length,
      writable: true
    },
    [Symbol.iterator]: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        return this;
      }
    },
    next: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        if (this.index < this.length) {
          return {
            value: array[this.index++],
            done: false
          };
        }
        return {
          done: true
        };
      }
    },
    return: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (value) {
        return {
          value: value,
          done: true
        };
      }
    },
    throw: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (exception) {
        return {
          done: true
        };
      }
    }
  });

  return Object.seal(iterator);
}

/**
 * Creates a restartable iterator from a provided request function. The iterator allows sequential access to results split by pageToken, and offers the ability to 'restart' iteration from the beginning.
 * @param {requestCallback} request 
 * @returns {Object} A sealed iterator object.
 */
function createRestartableIteratorFromPageTokenResponse_(request) {
  const iterator = new Object();

  Object.defineProperties(iterator, {
    index: {
      configurable: false,
      enumerable: true,
      value: undefined,
      writable: true
    },
    length: {
      configurable: false,
      enumerable: true,
      value: undefined,
      writable: true
    },
    [Symbol.iterator]: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        return this;
      }
    },
    next: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function () {
        const page = request(this.index);

        this.index = page.pageToken;

        if (this.index) {
          return {
            value: page,
            done: false
          };
        }
        return {
          value: page,
          done: true
        };
      }
    },
    return: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (value) {
        return {
          value: value,
          done: true
        };
      }
    },
    throw: {
      configurable: false,
      enumerable: true,
      writable: false,
      value: function (exception) {
        return {
          done: true
        };
      }
    }
  });

  return Object.seal(iterator);
}

/**
 * A custom error that can be thrown to stop the execution of the timeout function.
 * @class
 * @param {string} [message=''] - The message of the error.
 * @param {Object} [options] - The options object.
 * @param {Error} [options.cause] - The cause of the error.
 */
function StopError(message = '', options) {
  Object.defineProperties(this, {
    cause: {
      configurable: true,
      enumerable: false,
      value: undefined,
      writable: true
    }
  });
  this.message = message;
  if (options && options.cause) {
    this.cause = options.cause;
  }
}
StopError.prototype = Error.prototype;
StopError.prototype.name = 'StopError';
