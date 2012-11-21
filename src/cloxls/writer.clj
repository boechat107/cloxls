(ns cloxls.writer 
  (:import 
   [java.io IOException FileOutputStream]
   [org.apache.poi.hssf.usermodel HSSFWorkbook]
   )
  )


(defonce ^{:dynamic true
           :doc "This variable is bound to the created workbook when the macro
                  with-new-wb is used." }
  *wb* nil)


(defonce ^{:dynamic true
           :doc "This variable is bound to a sheet when the WITH-SHEET macro is used."}
  *sheet* nil)


(defmacro with-new-wb
  "Binds the variable *new-wb* to a new workbook whose name is given."
  [filename & body]
  `(let [file# (FileOutputStream. ~filename)]
     (binding [*wb* (HSSFWorkbook.)]
       ~@body
       (.write *wb* file#)
       (.close file#))))
;TODO: macro with-wb for a existing workbook.


(defn create-sheet!
  "Creates a sheet with the given name or with a default name. The integer index of the
   created sheet is returned." 
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! wb (str "Sheet " n) n)))
  ([wb sheet-name]
   (.createSheet wb sheet-name)
   (.getSheetIndex wb sheet-name)))


(defmacro with-sheet
  "Binds the variable *sheet* to a sheet with the given sheet-id of the workbook wb,
   exposes it to the body and write the modifications to file."
  [wb sheet-id & body]
  `(let [wb# ~wb]
     (binding [*sheet* (.getSheet wb# ~sheet-id)]
       ~@body)))
