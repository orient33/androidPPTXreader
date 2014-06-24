this is from https://github.com/qhm123/POI-Android pptx. part.
I modify some, to make it build with gradle .

if dx ERROR  with notice --core-library, you can modify dx yourself.
	as below:

```
cd  sdk/buildtools/*/
mv dx  dxbak
touch dx
vi dx 在dx中添加如下内容:
---------
#/bin/bash
/绝对路径/sdkbuildtoos/*/dxbak  --core-library  $@
---------
最后修改dx的权限  chmod +x dx即可

```
或者有更好的方法，比如在 build.gradle中添加 dexOptions{}的配置，




