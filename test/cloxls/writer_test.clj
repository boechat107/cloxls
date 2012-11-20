(ns cloxls.writer-test
    (:require
      [cloxls.writer :as w]))

(defn test1
  []
  (with-open [wb (w/create-workbook "test.xls")]
    (w/create-sheet! wb)
    (w/with-sheet wb 0
      (w/add-cell-data! 0 0 "Company")
      (w/add-cell-data! 0 1 "Number of Employees")
      (w/add-cell-data! 1 0 "Dog show")
      (w/add-cell-data! 1 1 10)
      )
    ))

(defn test2
  []
  (with-open [wb (w/create-workbook "test2.xls")]
    (w/create-sheet! wb)
    (w/with-sheet wb 0
      (w/add-line-data! 0 ["Company" "Number of Employees"])
      (w/add-line-data! 1 ["Dog show" 15])
      (w/add-line-data! 2 ["Lazy comp" 3])
      (w/add-line-data! 3 ["Total" "=B2+B3"])
      )))

(defn test3 
  [] 
  (with-open [wb (w/create-workbook "test3.xls")]
    (w/create-sheet! wb)
    (w/with-sheet wb 0
      (w/add-2d-data! [["Company" "Number of Employees"]
                       ["Dog show" 250]
                       ["Lazy comp" 5.0]
                       ["Total" "=SUM(B2,B3)"]]))))


