# cloxls

A Clojure library designed to write XLS files using the 
[Apache POI](http://poi.apache.org/) library.

## Installation

Add to the dependencies of the leiningen project:

```clj
[org.clojars.boechat107/cloxls "0.1.2-SNAPSHOT"]
```

## Usage

```clj
(use '(cloxls.core))

(defn writing-test
  []
  (with-new-wb "test_poi.xls"
    (create-sheet!)
    ;; The first created sheet has the default index 0.
    (with-sheet 0
      (create-2d-data! [["Company" "Number of employees"]
                        ["Dog show" 15]
                        ["Lazy comp" 30]
                        ["Total employees" "=B2+B3"]])
      (create-row-data! 4 ["Alternative" "function"])
      (create-row-data! 4 3 ["Additional" "information"])
      (create-row-data! 5 ["Form" "=B4+B7"])
      ;; Add data to the column E, starting at row 2.
      (create-col-data! 4 2 ["Column" "data"])
      (create-row-data! 6 ["Hidden cell" {:value 20 :hidden? true}])
      ;; Resize the columns' width to fit contents.
      (autosize-columns!))))

(defn reading-test
  []
  (with-wb "test_poi.xls"
    (sheet->matrix 0 true)))
```

See [docs](cloxls/blob/master/docs/index.html) for more details about the library functions.

## License

Copyright © 2012

Distributed under the Eclipse Public License, the same as Clojure.
