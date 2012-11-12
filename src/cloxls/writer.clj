(ns cloxls.writer
    "Functions to write data to a XLS file.
     References:
        http://www.vogella.com/articles/JavaExcel/article.html
        http://www.andykhan.com/jexcelapi/tutorial.html#writing"
    (:import
      [java.util Locale]
      [java.io File IOException]
      [jxl CellView Workbook WorkbookSettings]
      [jxl.write Formula Label WritableCellFormat WritableFont WritableSheet
                 WritableWorkbook WriteException]
      [jxl.write.biff RowsExceededException]
      )
    )

;; TODO: binding macro for a sheet or spreadsheet.

(defn create-workbook
  "Creates and returns the high level abstraction of the spreadsheet, a WritableWorkbook
   object. The default Locale is EN." 
  [filename]
  (let [file (File. filename)
        wb-set (WorkbookSettings.)]
    (.setLocale wb-set (Locale. "en" "EN"))
    (try
      (Workbook/createWorkbook file wb-set)
      (catch IOException e
             (str "Problem to create a xls file: " (.getMessage e))))))


(defn create-sheet!
  "Creates a sheet with the given name or with a default name. Side effects only."
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! wb (str "Sheet " n) n)))
  ([wb sheet-name idx]
   (.createSheet wb sheet-name idx)))

(defn add-label!
  "Add a text label to a specific cell."
  [sheet row col text]
  (.addCell sheet (Label. row col text)))

(defn add-number!
  "Add a number to a specific cell."
  [sheet row col num]
  (.addCell sheet (jxl.write.Number. row col num)))

(defn add-cell-data!
  "If the data is a string, a label is created. Otherwise, a number is created."
  [sheet row col data]
  (let [dtype (if (number? data)
                (jxl.write.Number. row col data)
                (Label. row col data))]
    (.addCell sheet dtype)))

(defn write-data!
  [wb]
  (doto wb (.write) (.close)))


