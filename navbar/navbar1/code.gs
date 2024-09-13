function doGet(e){   
    let page = e.parameter.page;
    if (page == null) page = "main";
    var output = HtmlService.createTemplateFromFile(page);
    return output.evaluate()
        .setTitle("template")
        .setFaviconUrl('https://cdn4.iconfinder.com/data/icons/social-media-logos-6/512/121-css3-512.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(){
    return HtmlService.createTemplateFromFile("header.html").evaluate().getContent();
}

function myURL(){
    return ScriptApp.getService().getUrl();
}
