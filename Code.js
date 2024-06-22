/**
 * The timeout() function executes the specified code on a set of values sourced from an iterable object. Sets a timer and restarts the specified code when the execution time approaches the Google Apps Script time limit.
 * @param {function} caller - The function to be executed.
 * @param {object} event - The event object that triggered the function.
 * @param {object} iterator - The iterable object that contains the values to be processed.
 * @param {function} callback - The function to be executed for each value in the iterable object.
 * @param {object} [options] - The options object.
 * @param {number} [options.timeout] - The maximum amount of time (in milliseconds) that the function can run before it is interrupted.
 * @param {number} [options.delay] - The amount of time (in milliseconds) to wait before restarting the function.
 * @param {number} [options.start] - The start time (in milliseconds) of the function.
 * @param {boolean} [options.debug] - Whether or not to enable debug logging.
 */
function timeout(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    // Default timeout of 5 minutes (in milliseconds)
    timeout: 5 * 60 * 1000,
    // Default delay of 1 minute (in milliseconds)
    delay: 1 * 60 * 1000,
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

  if (event) {
    const value = PropertiesService.getUserProperties()
      .getProperty(event.triggerUid);

    if (value === null) {
      throw new Error('Cannot find property for %s.', event.triggerUid);
    }

    const property = {
      index: iterator.index,
      length: iterator.length
    };

    Object.assign(property, JSON.parse(value));
    console.info('Restore property.', property);

    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === event.triggerUid) {
        ScriptApp.deleteTrigger(trigger);
        console.log('Delete trigger: %s', event.triggerUid);

        PropertiesService.getUserProperties()
          .deleteProperty(event.triggerUid);
        console.log('Delete property.');
      }
    }

    iterator.index = property.index;
    iterator.length = property.length;
  }

  for (const value of iterator) {
    const stop = callback(value);

    if (stop) {
      console.log('Callback function returnd [stop = true].');
      return;
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
      .setProperty(trigger.getUniqueId(), JSON.stringify(property));
    console.info('Save property.', property);
  }

  return;
}

/**
 * Execute the specified function in parallel with timeout handling.
 * @param {function} caller - The function to be executed.
 * @param {object} event - The event object that triggered the function.
 * @param {object} iterator - The iterable object that contains the values to be processed.
 * @param {function} callback - The function to be executed for each value in the iterable object.
 * @param {object} [options] - The options object.
 * @param {number} [options.split] - The number of times to split the iterable object.
 * @param {number} [options.timeout] - The maximum amount of time (in milliseconds) that the function can run before it is interrupted.
 * @param {number} [options.delay] - The amount of time (in milliseconds) to wait before restarting the function.
 * @param {number} [options.start] - The start time (in milliseconds) of the function.
 * @param {boolean} [options.debug] - Whether or not to enable debug logging.
 */
function timeoutInParallel(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    // Default timeout of 5 minutes (in milliseconds)
    timeout: 5 * 60 * 1000,
    // Default delay of 1 minute (in milliseconds)
    delay: 1 * 60 * 1000,
    // Record the start time
    start: Date.now(),
    // Default split value 
    split: 4,
    // Default debug logging to off
    debug: false
  }, options);

  if (options.split > 4) {
    throw new RangeError('options.split cannot be greater than 4.');
  }

  if (!event) {
    if (Math.pow(2, options.split) > (20 - ScriptApp.getProjectTriggers().length)) {
      throw new Error('Limit will be exceede.');
    }

    createSplitTrigger_(options.split, iterator.index, iterator.length, caller, options);

    return;
  }

  timeout(caller, event, iterator, callback, options);

  return;
}

/**
 * Recursively creates split triggers for time-based execution of a function. Designed to break down a larger task into smaller, scheduled executions.
 * @param {number} split - The number of splits remaining, used to determine recursion depth.
 * @param {number} index - The starting index for the current task segment.
 * @param {number} length - The total length of the task.
 * @param {function} caller - The function to be executed by the triggers.
 * @param {object} options - An object containing configuration options for the triggers:
 */
function createSplitTrigger_(split, index, length, caller, options) {
  const section = length - index;
  const left = Math.floor(section / 2);
  if (split > 0 && left > 1) {
    createSplitTrigger_(split - 1, index, index + left, caller, options);
    createSplitTrigger_(split - 1, index + left, length, caller, options);
  } else {
    const trigger = ScriptApp.newTrigger(caller.name)
      .timeBased()
      .after(options.delay)
      .create();
    console.info('Create trigger. Function name: %s Unique ID: %s', trigger.getHandlerFunction(), trigger.getUniqueId());

    const property = {
      index: index,
      length: length
    };

    PropertiesService.getUserProperties()
      .setProperty(trigger.getUniqueId(), JSON.stringify(property));
    console.info('Save property.', property);
  }
}

/**
 * Creates a restartable iterator from a provided range object. The iterator allows sequential access to values within the range, and offers the ability to 'restart' iteration from the beginning.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object to iterate over.
 * @returns {object} A sealed iterator object.
 */
function createRestartableIteratorFromRange(range) {
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
 * @param {object[][]} array - The 2D array to iterate over.
 * @returns {object} A sealed iterator object.
 */
function createRestartableIteratorFromArray(array) {
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
 * A callback function that calls the API to return results split by pageToken
 * @callback requestCallback
 * @param {string} pageToken
 * @returns {*}
 */

/**
 * Creates a restartable iterator from a provided request function. The iterator allows sequential access to results split by pageToken, and offers the ability to 'restart' iteration from the beginning.
 * @param {requestCallback} request 
 * @returns 
 */
function createRestartableIteratorFromPageTokenResponse(request) {
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
