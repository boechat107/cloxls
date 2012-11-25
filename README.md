# cloxls

A Clojure library designed to write XLS files using the 
[Apache POI](http://poi.apache.org/) library.

## Usage

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
          (create-row-data! 4 ["Alternative" "function"]))))
    
    (defn reading-test
      []
      (with-wb "test_poi.xls"
        (sheet->matrix 0 true)))

See [docs](docs/index.html) for more details about the library functions.

## License

Copyright © 2012

Distributed under the Eclipse Public License, the same as Clojure.
