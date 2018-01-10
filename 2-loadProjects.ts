import { JsomNode, IJsomNodeSettings } from 'sp-jsom-node';

var settings: any = require('./config/private.json');
declare const PS: any;

let jsomNodeOptions: IJsomNodeSettings = {
  siteUrl: settings.siteUrl,
  authOptions: <any>settings,
  modules: [ 'core', 'project' ]
};

(new JsomNode(jsomNodeOptions)).init();
//new JsomNode().wizard();

//Code here
var projContext;
var projects;

projContext = PS.ProjectContext.get_current();
projects = projContext.get_projects();

projContext.load(projects, 'Include(Name, Id)');

projContext.executeQueryAsync(
  //Success
  () => {

    var p, projId;
    var pTable = [];
    var pEnum = projects.getEnumerator();

    while (pEnum.moveNext()) {
      p = pEnum.get_current();

      var project = p;
      var projId = p.get_id();
      var projName = p.get_name();
      console.log(projId + ' ' + projName + '|');
    }
  },
  //Fail
  (sender, args) => {
    console.log("Failed: " + args.get_message());
    return;
  }
);
