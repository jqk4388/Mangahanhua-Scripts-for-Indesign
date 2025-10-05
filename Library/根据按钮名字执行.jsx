var action = app.menuActions.itemByName("复制选定链接的信息");
if (!action.isValid) action = app.menuActions.itemByName("Copy Link Info");
if (action.isValid) action.invoke();
else alert("找不到菜单命令");
