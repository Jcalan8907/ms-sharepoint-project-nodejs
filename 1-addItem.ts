import { JsomNode } from 'sp-jsom-node';
 
(new JsomNode()).wizard().then((settings) => {
 
  /// ... <<< JSOM can be used here
 
  let ctx = SP.ClientContext.get_current();
  var oList = ctx.get_web().get_lists().getByTitle('New Lists');

  var itemCreateInfo = new SP.ListItemCreationInformation();
    let oListItem = oList.addItem(itemCreateInfo);
        
    oListItem.set_item('Title', 'my record');
        
    oListItem.update();

    ctx.load(oListItem);
        
    ctx.executeQueryAsync(() => {
        console.log('added!');
      }, (sender, args) => {
        console.log(args.get_message());
      });
 
}).catch(console.log);