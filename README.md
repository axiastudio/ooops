ooops
=====

```java
  IWas.create()
    .load(new FileInputStream(INPUT_FILE))
    .offset(30f, 700f)
    .text("COMUNE", 10, 0f, 42f)
    .text("di xyz", 10, 0f, 34f)
    .text("Prot.N.", 8, 0f, 25f)
    .text("201300007719", 10, 0f, 16f)
    .text("04-04-13 08:11", 8, 0f, 8f)
    .text("f_728", 8, 0f, 0f)
    .datamatrix("c_f728#201300007719#673dfc1792", DatamatrixSize._22x22, 75f, 15f, 1.6f)
    .toStream(new FileOutputStream(OUTPUT_FILE));
```
