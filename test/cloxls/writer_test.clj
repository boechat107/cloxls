(ns cloxls.writer-test
  (:use
   [cloxls.writer]))


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


(defn test2
  []
  (with-new-wb "test_poi2.xls"
    (create-sheet!)
    (with-sheet 0
      (create-2d-data! [["Company" "Number of employees"]
                        ["Dog show" 10]
                        ["Lazy comp" 30]
                        ["Total employees" "=B2+B3"]]))))
