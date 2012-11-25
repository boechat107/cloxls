(defproject cloxls "0.1.1"
  :description "Clojure library to write and read XLS files."
  :url "https://github.com/boechat107/cloxls"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.4.0"]
                 [jxl "2.6.12"]
                 [poi "3.8"]]
  :profiles {:dev {:resource-paths ["resources"] :dependencies [[markdown-clj "0.9.10"]]}}
;  :dev-dependencies [[markdown-clj "0.9.10"]]
;  :resource-paths ["resources"]
  :html5-docs-ns-includes #"^cloxls\..*"
  :html5-docs-ns-excludes #".*jxl.*"
  :html5-docs-repository-url "https://github.com/boechat107/cloxls"
  )
