(ns cloxls.core
    "References:
        http://poi.apache.org/spreadsheet/quick-guide.html"
  (:import 
   [java.io IOException FileOutputStream FileInputStream]
   [org.apache.poi.hssf.usermodel HSSFWorkbook HSSFCell HSSFFormulaEvaluator]
   [org.apache.poi.poifs.filesystem POIFSFileSystem]
   )
  )


(defonce ^{:dynamic true
           :doc "This variable is bound to the created workbook when the macro
                  with-new-wb is used." }
  *wb* nil)


(defonce ^{:dynamic true
           :doc "This variable is bound to a sheet when the WITH-SHEET macro is used."}
  *sheet* nil)

;; ============== Workbook manipulation ==============================

(defmacro with-new-wb
  "Binds the variable *new-wb* to a new workbook whose name is given."
  [filename & body]
  `(let [file# (FileOutputStream. ~filename)]
     (binding [*wb* (HSSFWorkbook.)]
       ~@body
       (.write *wb* file#)
       (.close file#))))


(defmacro with-wb
  "Binds the variable  *new-wb* to a existing workbook whose name is given.
   Use this macro just for reading."
  [filename & body]
  `(let [file# (FileInputStream. ~filename)]
     (binding [*wb* (HSSFWorkbook. (POIFSFileSystem. file#))]
       (let [res# ~@body] 
         (.close file#)
         res#))))

;; ============== Sheet manipulation ==============================

(defn create-sheet!
  "Creates a sheet with the given name or with a default name. The integer index of the
   created sheet is returned."
  ([] (create-sheet! *wb*))
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! (str "Sheet " n) wb)))
  ([sheet-name wb]
   (.createSheet wb sheet-name)
   (.getSheetIndex wb sheet-name)))


(defmacro with-sheet
  "Binds the variable *sheet* to a sheet with the given sheet-id of the workbook *wb*,
   exposes it to the body and write the modifications to file. "
  [sheet-id & body]
  `(let [sheet-id# ~sheet-id]
     (binding [*sheet* (if (number? sheet-id#)
                         (.getSheetAt *wb* sheet-id#)
                         (.getSheet *wb* sheet-id#))]
      ~@body)))

;; ============== Adding cell contents ==============================

(defn create-cell!
  "Creates a cell, adds data to it and sets it on the sheet/workbook.
   The data could be a number or a string. If the string starts with =, a formula is
   created."
  [row-obj c-id data]
  (let [cell (.createCell row-obj c-id)]
    (cond
      (and (string? data)
           (= \= (get data 0))) (.setCellFormula cell (subs data 1))
      (number? data) (.setCellValue cell (double data))
      :else (.setCellValue cell data))))


(defn- coll-idx-data
  "Returns a new collection composed of tuples (idx, element), where idx is the index of
   the element in the original collection."
  [coll]
  (partition 2 (interleave (range (count coll)) coll)))


(defn create-row-data!
  "Creates a row of data from a data collection."
  ;; TODO: offset values.
  ([r-id data] (create-row-data! *sheet* r-id data))
  ([sheet r-id data]
     (let [row-obj (.createRow sheet r-id)]
       (doseq [[c-id d] (coll-idx-data data)]
         (create-cell! row-obj c-id d)))))


(defn create-2d-data!
  "Adds the data to a sheet as a matrix, starting from line 0 and column 0."
  ([data] (create-2d-data! *sheet* data))
  ([sheet data]
     (doseq [[r-id ld] (coll-idx-data data)]
       (create-row-data! sheet r-id ld))))


;; ============== Reading cell contents ==============================

(defn get-cell-content
  "Gets the content of the cell considering its type. If formula-eval? is true, the cell 
   is evaluated and formulas results are gotten."
  ;; TODO: support date format.
  ([cell formula-eval?] (get-cell-content *wb* cell formula-eval?))
  ([wb cell formula-eval?]
   (let [eval-type (when formula-eval?
                     (-> (.getCreationHelper wb)
                         (.createFormulaEvaluator)
                         (.evaluateFormulaCell cell)))
         cell-type (if (or (nil? eval-type) (= -1 eval-type))
                     (.getCellType cell)
                     eval-type)]
     (condp = cell-type
       HSSFCell/CELL_TYPE_STRING (-> (.getRichStringCellValue cell)
                                     (.getString))
       HSSFCell/CELL_TYPE_NUMERIC (.getNumericCellValue cell)
       HSSFCell/CELL_TYPE_BOOLEAN (.getBooleanCellValue cell)
       HSSFCell/CELL_TYPE_FORMULA (.getCellFormula cell)))))


(defn sheet->matrix
  "Gets the contents of specific sheet of a workbook. Formulas are gotten as calculated 
   values if it is possible.
   The sheet-id must be a integer (index) or a string (name).
   If formula-eval? is true, the cells are evaluated and formulas results are gotten
   (default false)."
  ([sheet-id] (sheet->matrix *wb* sheet-id nil))
  ([sheet-id formula-eval?] (sheet->matrix *wb* sheet-id formula-eval?))
  ([wb sheet-id formula-eval?]
     (let [sheet (if (number? sheet-id)
                   (.getSheetAt wb sheet-id)
                   (.getSheet wb sheet-id))]
       (vec (map (fn [row-obj]
                   (vec (map #(get-cell-content wb % formula-eval?)
                             (iterator-seq (.cellIterator row-obj)))))
                 (iterator-seq (.rowIterator sheet)))))))
