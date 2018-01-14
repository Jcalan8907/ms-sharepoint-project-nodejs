import { JsomNode, IJsomNodeSettings } from 'sp-jsom-node';

var settings: any = require('./config/private_macbook.json');
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

var lookupTables = context.get_lookupTables();

context.load(lookupTables);

var myLut = lookupTables.getByGuid('ebb134c6-29be-e711-80d3-00155da4760e');

var myEntries = myLut.get_entries();
context.load(myEntries);

var newId = SP.Guid.newGuid();

var lutEntry = new PS.LookupEntryCreationInformation();

lutEntry.set_description("my descr");
lutEntry.set_sortIndex(10); //index of row to sort. How to learn count of rows?
lutEntry.set_id(newId);
lutEntry.set_parentId(null);

var lutValue = new PS.LookupEntryValue();
lutValue.set_textValue("test");
lutEntry.set_value(lutValue);
myEntries.add(lutEntry);

lookupTables.update(myEntries);

context.executeQueryAsync(

    
    () => {
      //Success

    },
    
    (sender, args) => {
      //Fail

      console.log("Failed: " + args.get_message());
      return;
      
    }
);