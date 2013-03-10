(ns cloxls.writer-test
  (:use
   [cloxls.core]))


(defn test1
  []
  (with-new-wb "test_poi1.xls"
    ;; Creates a sheet whose index is 0.
    (create-sheet!)
    (with-sheet 0
      (create-row-data! 0 ["Company" "Number of employees"])
      (create-row-data! 1 ["Dog show" 10])
      (create-row-data! 2 ["Lazy comp" 30])
      (create-row-data! 3 ["Total employees" "=B2+B3"]))))


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
      ;; Add data to the column E.
      (create-col-data! 4 2 ["Column" "data"])
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
