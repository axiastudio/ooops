ooops
=====

```java
List<List<String>> books = new ArrayList<>();
books.add(null); // skip first line
books.add(new ArrayList<>(Arrays.asList("Anna Karenina", "Tolstoy")));
books.add(new ArrayList<>(Arrays.asList("The Master and Margarita", "Bulgakov")));

Ooops.create()
        .open("uno:socket,host=localhost,port=8997;urp;StarOffice.ServiceManager")
        .load(new FileInputStream(new File("test.odt")))
        .fillBookmark("name", "Tiziano")
        .fillTable("bookTable", books)
        .showHideSection("appear", true)
        .showHideSection("notAppear", false)
        .filter(Filter.PDF)
        .toStream(new FileOutputStream(new File("test.pdf")));
```

You need an [Open|Libre]Office listener.

For example (Docker):

```bash
docker run -p 8997:8997 -d xcgd/libreoffice
```