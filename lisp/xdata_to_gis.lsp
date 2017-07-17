(defun xdata_to_gis ( / e i n s x l u b f r)
  ;;; This function creates gis content from autocad xdata
  ;;; by using the ArcGIS for AutoCAD API functions
  (vl-load-com)
  (esri_featureclass_add  "SPACEDATA"
                         (list
                           (cons "GEOMTYPE" "POLYGON")
                           (cons "LAYERFILTER" "A-SPAC-PPLN-AREA")
                           ))
  
  (esri_fielddef_add "SPACEDATA" 
                     (list 
                       (cons "name" "ROOM_ID") 
                       (cons "Type"  "String") 
                       (cons "Length" 30)
                       )
                     )
  (if (setq s (ssget "x" (list (cons 8 "A-SPAC-PPLN-AREA"))))
      (progn
        (setq i 0
              n (sslength s)
              )
        (while (< i n)
               (setq e (ssname s i)        
                     i (1+ i)
                     )
               (if (= (type e) 'ENAME)
                   (progn 
                     (setq u (vlax-ename->vla-object e)   
                           ;;; this line captures the APLS's xdata which is the most
                           ;;; reliable source of room ids across our drawings               
                           v (vla-getxdata u "APLS_FM" 'codes 'values)
                           )
                     (if (and codes
                              values
                              )
                         (progn
                           (setq  room_id (nth 0 (cdr (mapcar 'variant-value (vlax-safearray->list values))))
                                 dwg_name  (strcat  (vl-filename-base (getvar "dwgname")))
                                 bldg_floor (vl-string-trim "DWG" (vl-string-trim "-BAS" dwg_name))
                                 ;;;getting unique id i.e. UCB-302-01-CRN160
                                 id (strcat "UCB" "-" bldg_floor room_id)
                                 )
                           (esri_attributes_set
                             e
                             "SPACEDATA"
                             (list 
                               (cons "ROOM_ID" id)
                               );list       
                             );esri
                           (print room_id)
                           );;;progn
                         );;;if                  
                     );progn
                   );if             
               );;;while
        )
      )
  (princ)
  )
(xdata_to_gis)
