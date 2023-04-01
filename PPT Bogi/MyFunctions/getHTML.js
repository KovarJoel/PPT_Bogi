function getHTML() {

    try {
        //return document.querySelector('html')[0].innerHTML;
        //return document.getElementsByTagName('html')[0].innerHTML;
        //return document.documentElement.innerHTML;
        //return document.body.parentNode.innerHTML;
        //return document.all[0].innerHTML;
        //return document.head.parentNode.innerHTML;

        return document.documentElement.outerHTML;
    }
    catch (err) {
        return err.message;
    }
    
}

getHTML();