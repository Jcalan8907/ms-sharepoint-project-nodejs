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

var context = PS.ProjectContext.get_current();
// Retrieve custom fields and lookup table definitions
var custom_fields = context.get_customFields();
var jur_custom_field = custom_fields.getByGuid('09e501e7-29be-e711-80ea-00155da4370e');

var project_lookup_table_entries = jur_custom_field.get_lookupEntries();

context.load(jur_custom_field);

context.executeQueryAsync(

    
    () => {
      //Success

      var n = jur_custom_field.get_name();
      console.log(n);

    },
    
    (sender, args) => {
      //Fail

      console.log("Failed: " + args.get_message());
      return;
      
    }
);


