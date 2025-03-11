/**
 * @file The library provides a way to set a timeout on a function call: if the function does not complete within the specified timeout, a timer is set to resume processing after a certain amount of time.
 */

// Use Strict mode.
'use strict';

// Default timeout of 5 minutes (in milliseconds)
const DEFAULT_TIMEOUT = 300000; // 5 * 60 * 1000
// Default delay of 1 minute (in milliseconds)
const DEFAULT_DELAY = 60000; // 1 * 60 * 1000
// Default split value 
const DEFAULT_SPLIT = 4;
// Default debug logging to off
const DEFAULT_DEBUG = false;

/**
 * The timeout() function executes the specified code on a set of values sourced from an iterable object. Sets a timer and restarts the specified code when the execution time approaches the Google Apps Script time limit.
 * @param {(event: GoogleAppsScript.Events.TimeDriven) => void} caller - The function to be executed.
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 * @param {(Array<any>|GoogleAppsScript.Spreadsheet.Range|(pageToken: string) => any)} iterator - A union containing the values ​​to be processed.
 * @param {(value: any) => void} callback - A function that is executed for each value in the iterable object, and will stop processing if it raises a StopError.
 * @param {{timeout?: number, delay?: number, start?: number, debug?: boolean}} [options] - The options object.
 * @returns {(GoogleAppsScript.Script.Trigger|null)} - a GoogleAppsScript.Script.Trigger or null
 */
function timeout(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    timeout: DEFAULT_TIMEOUT,
    delay: DEFAULT_DELAY,
    start: Date.now(),
    debug: DEFAULT_DEBUG
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

  if (Array.isArray(iterator)) {
    iterator = new RestartableIteratorFromArray(iterator);
  } else if (iterator.toString() === 'Range') {
    iterator = new RestartableIteratorFromRange(iterator);
  } else if (typeof iterator === 'function') {
    iterator = new RestartableIteratorFromRequestWithPageToken(iterator);
  }

  return timeout_(caller, event, iterator, callback, options);
}

/**
 * @param {(event: GoogleAppsScript.Events.TimeDriven) => void} caller - The function to be executed.
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 * @param {Object} iterator - An iterator object.
 * @param {(value: any) => void} callback - A function that is executed for each value in the iterable object, and will stop processing if it raises a StopError.
 * @param {{split?: number, timeout: number, delay: number, start: number, debug: boolean}} options - The options object.
 * @returns {(GoogleAppsScript.Script.Trigger|null)} - a GoogleAppsScript.Script.Trigger or null
 */
function timeout_(caller, event, iterator, callback, options) {
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
  }

  if ((!event || !event.triggerUid) ||
    (iterator !== undefined && iterator.length === undefined) ||
    (iterator.index < iterator.length)) {

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
 * @param {(event: GoogleAppsScript.Events.TimeDriven) => void} caller - The function to be executed.
 * @param {GoogleAppsScript.Events.TimeDriven} event - The event object that triggered the function.
 * @param {(Array<any>|GoogleAppsScript.Spreadsheet.Range)} iterator - A union containing the values ​​to be processed.
 * @param {(value: any) => void} callback - A function that is executed for each value in the iterable object, and will stop processing if it raises a StopError.
 * @param {{split?: number, timeout?: number, delay?: number, start?: number, debug?: boolean}} [options] - The options object.
 */
function timeoutInParallel(caller, event, iterator, callback, options) {
  // Merge default options with user provided options
  options = Object.assign({
    split: DEFAULT_SPLIT,
    timeout: DEFAULT_TIMEOUT,
    delay: DEFAULT_DELAY,
    start: Date.now(),
    debug: DEFAULT_DEBUG
  }, options);

  if (!iterator) {
    throw new ReferenceError('iterator is not defined.');
  }
  // Check if the iterator is a request object.
  if (typeof iterator === 'function') {
    throw new TypeError('iterators cannot be request.');
  }

  if (Array.isArray(iterator)) {
    iterator = new RestartableIteratorFromArray(iterator);
  } else if (iterator.toString() === 'Range') {
    iterator = new RestartableIteratorFromRange(iterator);
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
  } else {
    // If an event is provided, handle the timeout logic
    timeout_(caller, event, iterator, callback, options);
  }
}

/**
 * Recursively creates split triggers for time-based execution of a function. Designed to break down a larger task into smaller, scheduled executions.
 * @param {number} split - The number of splits remaining, used to determine recursion depth.
 * @param {number} index - The starting index for the current task segment.
 * @param {number} length - The total length of the task.
 * @param {(event: GoogleAppsScript.Events.TimeDriven) => void} caller - The function to be executed.
 * @param {{split?: number, timeout?: number, delay?: number, start?: number, debug?: boolean}} [options] - The options object.
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
 * A base class for creating restartable iterators.
 * @constructor
 */
function RestartableIterator_() { }
RestartableIterator_.prototype[Symbol.iterator] = function () {
  return this;
};
RestartableIterator_.prototype.return = function (value) {
  return {
    value: value,
    done: true
  };
};
RestartableIterator_.prototype.throw = function (exception) {
  return {
    done: true
  };
};
RestartableIterator_.prototype.next = undefined;
RestartableIterator_.prototype.index = undefined;
RestartableIterator_.prototype.length = undefined;

/**
 * Creates a restartable iterator from a provided array. The iterator allows sequential access to the inner arrays within the main array, and offers the ability to 'restart' iteration from the beginning.
 * @constructor
 * @param {Array<any>} array - The array to iterate over.
 */
function RestartableIteratorFromArray(array) {
  RestartableIterator_.call(this);
  this.array = array;
  this.index = 0;
  this.length = array.length;
}
RestartableIteratorFromArray.prototype = Object.create(RestartableIterator_.prototype, {
  constructor: {
    value: RestartableIteratorFromArray,
    enumerable: false,
    writable: true,
    configurable: true,
  },
  next: {
    value: function (value) {
      if (this.index < this.length) {
        return {
          value: this.array[this.index++],
          done: false
        };
      }

      return {
        done: true
      };
    },
    enumerable: false,
    writable: true,
    configurable: true,
  }
});

/**
 * Creates a restartable iterator from a provided range object. The iterator allows sequential access to values within the range, and offers the ability to 'restart' iteration from the beginning.
 * @constructor
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object to iterate over.
 */
function RestartableIteratorFromRange(range) {
  RestartableIterator_.call(this);
  this.range = range;
  this.index = 0;
  this.length = range.getNumRows();
}
RestartableIteratorFromRange.prototype = Object.create(RestartableIterator_.prototype, {
  constructor: {
    value: RestartableIteratorFromRange,
    enumerable: false,
    writable: true,
    configurable: true,
  },
  next: {
    value: function (value) {
      if (this.index < this.length) {
        return {
          value: this.range.offset(this.index++, 0, 1),
          done: false
        };
      }

      return {
        done: true
      };
    },
    enumerable: false,
    writable: true,
    configurable: true,
  }
});

/**
 * Creates a restartable iterator from a provided request function. The iterator allows sequential access to results split by pageToken, and offers the ability to 'restart' iteration from the beginning.
 * @constructor
 * @param {(pageToken?: string) => Object} request - The request function to iterate over.
 */
function RestartableIteratorFromRequestWithPageToken(request) {
  RestartableIterator_.call(this);
  this.request = request;
  this.index = undefined;
  this.length = undefined;
}
RestartableIteratorFromRequestWithPageToken.prototype = Object.create(RestartableIterator_.prototype, {
  constructor: {
    value: RestartableIteratorFromRequestWithPageToken,
    enumerable: false,
    writable: true,
    configurable: true,
  },
  next: {
    value: function (value) {
      const page = this.request(this.index);
      const done = page.pageToken ? false : true;

      this.index = page.pageToken;

      return {
        value: page,
        done: done
      };
    },
    enumerable: false,
    writable: true,
    configurable: true,
  }
});

/**
 * A custom error that can be thrown to stop the execution of the timeout function.
 * @class
 * @param {string} [message=''] - The message of the error.
 * @param {{cause: Error}} [options] - The options object. The cause of the error.
 */
function StopError(message = '', options) {
  Error.call(this, message, options);
  this.message = message;
  if (options && options.cause) {
    this.cause = options.cause;
  }
}
StopError.prototype = Object.create(Error.prototype, {
  constructor: {
    value: StopError,
    enumerable: false,
    writable: true,
    configurable: true,
  },
  name: {
    value: StopError.name,
    enumerable: false,
    writable: true,
    configurable: true,
  }
});
