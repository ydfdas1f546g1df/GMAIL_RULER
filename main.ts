type Criteria = {
    field: string;
    value: string;
};

enum Action {
    LABEL = 'label',
    DELETE = 'delete',
    MARK_READ = 'markRead',

}

function addNewRule(
    criteria: Array<Criteria>, action: Action, value: String, enabled: boolean, priority: number
):
    void {
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
    let rules = JSON.parse(scriptProperties.getProperty('rules') || '[]');
    let id = 0

    if (rules.length > 0) {
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

function testPushRules():void {
    addNewRule(
        [
            {
                "field": "from",
                "value": "*@example.com"
            }
        ],
        Action.LABEL,
        "Promotions",
        true,
        1
    );
  getRules();
}

function getRules() {
    let scriptProperties = PropertiesService.getScriptProperties();
    let rules = JSON.parse(scriptProperties.getProperty('rules') || '[]');
    console.log(rules);
}

function removeRules() {
    let scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('rules', "");
}

function createTimeDrivenTriggers() {
    ScriptApp.newTrigger('applyRules')
        .timeBased()
        .everyMinutes(1)
        .create();
}