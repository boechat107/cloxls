# cloxls

A Clojure library designed to write XLS files using the 
[Apache POI](http://poi.apache.org/) library.

## Installation

Add the last version of **cloxls** to the dependencies of the leiningen project:

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
      ;; Sets the font's size of A1.
      (set-font-style! 0 0 :size 30)
      ;; Conditional formatting: change manually the values of the cells composing
      ;; the rule and see what happens! :)
      (conditional-formatting! ["A4:B4" "A1:B1"]
                               [{:rule "$B$2>10", :font {:color :green}}
                                ;; Using a RGB similar color.
                                {:rule "$B$2<=10", :font {:color [150 0 50]}}])
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
* Change the font's style of cells.

## License

Copyright © 2012 Andre A. Boechat
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this
file except in compliance with the License. You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
