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


(def ^{:dynamic true
       :doc "This variable is bound to a sheet when the WITH-SHEET macro is used."}
       *sheet* nil)


(defn create-sheet!
  "Creates a sheet with the given name or with a default name and points the variable
   *sheet* to the new sheet. Side effects only." 
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! wb (str "Sheet " n) n)))
  ([wb sheet-name idx]
   (let [*sheet* (.createSheet wb sheet-name idx)]
     *sheet*)))


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
  ([row col data] (add-cell-data! *sheet* row col data))
  ([sheet row col data]
   (let [dtype (if (number? data)
                 (jxl.write.Number. row col data)
                 (Label. row col data))]
     (.addCell sheet dtype))))


(defmacro with-sheet
  "Binds the variable *sheet* to a sheet with the given sheet-id of the workbook wb,
   exposes it to the body and write the modifications to file."
  [wb sheet-id & body]
  `(let [wb# ~wb]
     (binding [*sheet* (.getSheet wb# ~sheet-id)]
       ~@body
       (.write wb#))))


(defn close-workbook!
  [wb]
  (doto wb (.write) (.close)))


