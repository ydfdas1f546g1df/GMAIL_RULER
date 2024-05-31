function addNewRule(criteria, action, value, enabled, priority) {
  /**
{
  "id": "000",
  "criteria": [
    {
      "field": "from",
      "value": "newsletter@example.com"
    },
    {
      "field": "subjectContains",
      "value": "special offer"
    }
  ],
  "action": "label",
  "value": "Promotions",
  "enabled": true,
  "priority": 1
}
  */

  let scriptProperties = PropertiesService.getScriptProperties();
  var rules = JSON.parse(scriptProperties.getProperty('rules') || '[]');
  var id = 0

  if (!(rules == [])) {
    id = rules[-1].id + 1
  }



  var ruleNew = {
    "id": id,
    "criteria": criteria,
    "action": action,
    "value": value,
    "enabled": enabled,
    "priority": priority,
  }

  rules.push(ruleNew);
  scriptProperties.setProperty('rules', JSON.stringify(rules));
}

function testPushRules(){
  addNewRule(
    [
      {
        "field": "from",
        "value": "*@example.com"
      }
    ],
    
  )
}

function getRules() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var rules = JSON.parse(scriptProperties.getProperty('rules') || '[]');
  console.log(rules);
}

function removeRules() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('rules', "");
}

function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('applyRules')
    .timeBased()
    .everyMinutes(1)
    .create();
}