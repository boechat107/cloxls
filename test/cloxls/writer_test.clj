(ns cloxls.writer-test
    (:require
      [cloxls.writer :as w]))

;(let [wb (w/create-workbook "test.xls")
;      sheet (w/create-sheet! wb)]
;  (doto sheet
;        (w/add-cell-data! 0 0 "Company")
;        (w/add-cell-data! 0 1 "Number of Employees")
;        (w/add-cell-data! 1 0 "Dog show")
;        (w/add-cell-data! 1 1 4)
;        )
;  (w/close-workbook! wb)
;  )

(with-open [wb (w/create-workbook "test.xls")]
  (w/create-sheet! wb)
  (w/with-sheet wb 0
    (w/add-cell-data! 0 0 "Company")
    (w/add-cell-data! 0 1 "Number of Employees")
    (w/add-cell-data! 1 0 "Dog show")
    (w/add-cell-data! 1 1 10)
    )
  )
