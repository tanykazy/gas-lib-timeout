/**
 * @constructor
 * @param {function(Object):void} handler
 * @param {Object} event
 * @param {Object} [options]
 * @param {number} [options.timeout = 5 * 60 * 1000]
 * @param {number} [options.delay = 1 * 60 * 1000]
 * @param {boolean} [options.debug = false]
 */
function Timeout(handler, event, options = {
  timeout: 5 * 60 * 1000,
  delay: 1 * 60 * 1000,
  debug: false
}) {
  this.start = Date.now();

  if (typeof handler !== 'function') {
    throw `handler is not function.`;
  }
  if (!handler.name) {
    throw `handler.name is not defined.`;
  }

  this.triggerBuilder = ScriptApp.newTrigger(handler.name)
    .timeBased()
    .after(options.delay);

  let property = null;
  if (event) {
    property = PropertiesService.getUserProperties()
      .getProperty(event.triggerUid);

    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === event.triggerUid) {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Delete trigger: ${event.triggerUid}`);

        PropertiesService.getUserProperties()
          .deleteProperty(event.triggerUid);
        console.log(`Delete context.`);
      }
    }
  }

  this.options = options;

  this.context = Object.assign({
    index: 0
  }, JSON.parse(property));

  return this;
}

/**
 * @param {function(void):Object[][]} initializer
 */
Timeout.prototype.initialize = function (initializer) {
  if (typeof initializer !== 'function') {
    throw `initializer is not function.`;
  }

  this.initializer = initializer;

  return this;
};

/**
 * @param {function(Object[], number, Object[][]):any} iterater
 */
Timeout.prototype.iterate = function (iterater) {
  if (typeof iterater !== 'function') {
    throw `iterater is not function.`;
  }

  this.iterater = iterater;

  return this;
};

/**
 * 
 */
Timeout.prototype.run = function () {
  const table = this.initializer();

  while (this.context.index < table.length) {
    const result = this.iterater(table[this.context.index], this.context.index, table);
    // console.log(`Callback Result: ${result}`);

    this.context.index++;

    const progress = Date.now() - this.start;
    if (progress > this.options.timeout) {
      if (this.context.index < table.length) {
        const trigger = this.triggerBuilder.create();
        console.log(`Create trigger. Function name: ${trigger.getHandlerFunction()} Unique ID: ${trigger.getUniqueId()}`);

        PropertiesService.getUserProperties()
          .setProperty(trigger.getUniqueId(), JSON.stringify(this.context));
        console.log(`Save context.`, this.context);

        return trigger;
      }
    }
  }

  return null;
};

function deleteAllTrigger_() {
  ScriptApp.getProjectTriggers().forEach(
    trigger => ScriptApp.deleteTrigger(trigger)
  )
}
