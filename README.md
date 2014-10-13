ooops
=====

```java
Map<String, Object> map = new HashMap<>();
map.put("name", "Tiziano");

Ooops.create()
  .open("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager")
  .load(new FileInputStream(new File("test.odt")))
  .map(map)
  .filter(Filter.writer_pdf_Export)
  .toStream(new FileOutputStream(new File("test.pdf")));
```
