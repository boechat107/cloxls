# cloxls

A Clojure library designed to write XLS files using the 
[Apache POI](http://poi.apache.org/) library.

## Installation

Add to the dependencies of the leiningen project:

```clj
[org.clojars.boechat107/cloxls "0.2.1-SNAPSHOT"]
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
      ;; Conditional formatting: change manually the values of the cells composing
      ;; the rule and see what happens! :)
      (conditional-formatting! ["A4:B4" "A1:B1"]
                               [{:rule "$B$2>10", :font {:color :green}}
                                ;; Using RGB values.
                                {:rule "$B$2<=10", :font {:color {:r 0 :g 0 :b 200}}}])
      ;; Resize the columns' width to fit contents.
      (autosize-columns!))))

(defn reading-test
  []
  (with-wb "test_poi.xls"
    (sheet->matrix 0 true)))
```

See [docs](http://boechat107.github.com/cloxls) for more details about the library functions.

## Features

* A new XLS file can be created;
* Data is inputed as a row, as a column or as a matrix;
* The type of the data can be number, text (a label) or a formula;
* The columns' width can be auto resized;
* Conditional formatting can be defined (for now, only font colors is modifiable);
* XLS files can be read and its contents returned as clojure vectors;
* The formulas of a XLS file can be evaluated before a reading, getting the result values 
instead of the formula definition.

## License

Copyright © 2012

Distributed under the Eclipse Public License, the same as Clojure.
