rm(list=ls())
library(RSelenium)
library(XML)
library(xlsx)

# iniciar navegador Chrome
rD <- rsDriver()
remDr <- rD[["client"]]

remDr$navigate(paste0("http://www.socioempleo.gob.ec/socioEmpleo-war/paginas/procesos/busquedaOferta.jsf?tipo=PRV"))
Sys.sleep(runif(1, 4, 5))#se espera un tiempo de hasta 5 segundos sin realizar nada hasta que la página se cargue por completo
doc1 <- htmlParse(remDr$getPageSource()[[1]]) #toma a la pagina activa y coge todo el código fuente de la pagina en la que se encuentra 
pag  <- doc1['//*[@id="anunciosIndex"]/table/tbody/tr/td[6]']#copiar el xpath desde chrome
npag <- try(toString(xmlValue(pag[[1]][[1]])))#el archivo es del tipo nodo, entonces hay que cambiarlo para poder leer el valor
npag <- sub(".*:", "", npag)
npag <- as.numeric(npag)
#npag <- npag - 1

num <- try(doc1['//*[@id="anunciosIndex"]/table/tbody/tr/td[4]'])#inpseccionar numero de registros, copiar xpath
nnum <- try(toString(xmlValue(num[[1]][[1]])))#casi siempre la informacion que se necesita esta en el nodo [[1]][[1]]
nnum <- sub("Total:", "", nnum)
nnum <- sub(" registros,", "", nnum)
nnum <- as.numeric(nnum)
nnum <- nnum

base <-matrix(nrow=nnum,ncol=17)


#creacion de las variables de la matriz 
variables <- c("cargo", "tipo", "relacion", "sector", "ciudad", "parroquia", "fechai", "fechaf", "instruccion", "remuneracion", "experiencia",
               "area", "conocimiento", "actividades", "jornada", "capacitacion",
               "adicional", "vacantes", "empresa")

for (i in 1:19){
  assign(paste0(letters[i],letters[i]), 1)
  assign(paste0(variables[i], "2"), matrix(, nrow = nnum, ncol = 1))
}

k <- 1
y <- 1
for (y in 1:(npag)){
  #for (y in 1:2){
  tryCatch({
    Sys.sleep(runif(1, 2, 4))# elegir randomicamente un número que sea de 6 a 8 segundo
    nclick <- y
    if (nclick!=1){
      click<-y-1
      remDr$findElement(using = "xpath", paste0("//*[@id='formBuscaOferta:pagina']/div[3]/span"))$clickElement()
      Sys.sleep(runif(1, 0.1, 0.3))
      remDr$findElement(using = "xpath",paste0("//*[@id='formBuscaOferta:pagina_",click,"']"))$clickElement()
    }
    Sys.sleep(runif(1,2,3))
    doc1 <- htmlParse(remDr$getPageSource()[[1]]) 
    
   
    ##este codigo se hace porque no funciona para todos los casos el codigo 51, pues cuando es importante, empresa es 61
    
    #lista de xpath de los links de las ofertas de trabajo
    clase <- c ("//*[@id='formBuscaOferta:listResult:0:j_idt43']/div/a",
                "//*[@id='formBuscaOferta:listResult:0:j_idt53']/div/a",
                "//*[@id='formBuscaOferta:listResult:1:j_idt43']/div/a",
                "//*[@id='formBuscaOferta:listResult:1:j_idt53']/div/a",
                "//*[@id='formBuscaOferta:listResult:2:j_idt43']/div/a",
                "//*[@id='formBuscaOferta:listResult:2:j_idt53']/div/a",
                "//*[@id='formBuscaOferta:listResult:3:j_idt43']/div/a",
                "//*[@id='formBuscaOferta:listResult:3:j_idt53']/div/a",
                "//*[@id='formBuscaOferta:listResult:4:j_idt43']/div/a",
                "//*[@id='formBuscaOferta:listResult:4:j_idt53']/div/a"
    )
    
    #lista de xpath de empresa
    clase2 <- c ("//*[@id='formBuscaOferta:listResult:0:j_idt43']/legend",
                 "//*[@id='formBuscaOferta:listResult:0:j_idt53']/legend",
                 "//*[@id='formBuscaOferta:listResult:1:j_idt43']/legend",
                 "//*[@id='formBuscaOferta:listResult:1:j_idt53']/legend",
                 "//*[@id='formBuscaOferta:listResult:2:j_idt43']/legend",
                 "//*[@id='formBuscaOferta:listResult:2:j_idt53']/legend",
                 "//*[@id='formBuscaOferta:listResult:3:j_idt43']/legend",
                 "//*[@id='formBuscaOferta:listResult:3:j_idt53']/legend",
                 "//*[@id='formBuscaOferta:listResult:4:j_idt43']/legend",
                 "//*[@id='formBuscaOferta:listResult:4:j_idt53']/legend"
    )
    
    
    for (i in clase2){
      empresa <-  try(doc1[i]) 
      if (!is.null(empresa)){
        empresa2[ss,1] <- try(toString(xmlValue(empresa[[1]][[1]])))
        ss <- ss+1
      }else {}
    }
    for (j in clase){
      doc1 <- htmlParse(remDr$getPageSource()[[1]])
      Sys.sleep(runif(1, 0.3, 0.5))
      link <- try(doc1[j])
      if (!is.null(link)){
        remDr$findElement(using = "xpath", paste0(j))$clickElement()
        Sys.sleep(runif (1,0.1,0.2))
        page_source <- remDr$getPageSource()
        doc <- htmlParse(remDr$getPageSource()[[1]])
       
        Sys.sleep(runif (1,0.1,0.2))
        
        cargo <- try(doc['//*[@id="formBuscaOferta:olGrid"]/tbody/tr[1]/td[2]']) 
        cargo2[aa,1] <- try(toString(xmlValue(cargo[[1]][[1]])))
        base[k,1] <- cargo2[aa,1] 
        aa <- aa+1
        Sys.sleep(runif (1,0.1,0.2))
        
        tipo <- try(doc['//*[@id="formBuscaOferta:j_idt36_label"]']) 
        tipo2[bb,1] <- try(toString(xmlValue(tipo[[1]][[1]])))
        base[k,2] <- tipo2[bb,1] 
        bb <- bb+1
        Sys.sleep(runif (1,0.1,0.2))
        
        relacion <- try(doc['//*[@id="formBuscaOferta:j_idt39_label"]']) 
        relacion2[cc,1] <- try(toString(xmlValue(relacion[[1]][[1]])))
        base[k,3] <- relacion2[cc,1] 
        cc <- cc+1
        Sys.sleep(runif (1,0.1,0.2))
        
        sector <- try(doc['//*[@id="formBuscaOferta:j_idt51_label"]']) 
        sector2[dd,1] <- try(toString(xmlValue(sector[[1]][[1]])))
        base[k,4] <- sector2[dd,1] 
        dd <- dd+1
        Sys.sleep(runif (1,0.1,0.2))
        
        ciudad <- try(doc['//*[@id="formBuscaOferta:j_idt47"]']) 
        ciudad2[ee,1] <- try(toString(xmlValue(ciudad[[1]][[1]])))
        base[k,5] <- ciudad2[ee,1] 
        ee <- ee+1
        Sys.sleep(runif (1,0.1,0.2))
        
        parroquia <- try(doc['//*[@id="formBuscaOferta:j_idt49"]']) 
        parroquia2[ff,1] <- try(toString(xmlValue(parroquia[[1]][[1]])))
        base[k,6] <- parroquia2[ff,1]
        ff <- ff+1
        Sys.sleep(runif (1,0.1,0.2))
        
        fechai <- try(doc['//*[@id="formBuscaOferta:olGrid"]/tbody/tr[5]/td[2]']) 
        fechai2[gg,1] <- try(toString(xmlValue(fechai[[1]][[1]])))
        base[k,7] <- fechai2[gg,1] 
        gg <- gg+1
        Sys.sleep(runif (1,0.1,0.2))
        
        fechaf <- try(doc['//*[@id="formBuscaOferta:olGrid"]/tbody/tr[5]/td[4]']) 
        fechaf2[hh,1] <- try(toString(xmlValue(fechaf[[1]][[1]])))
        base[k,8] <- fechaf2[hh,1] 
        hh <- hh+1
        Sys.sleep(runif (1,0.1,0.2))
        
        instruccion <- try(doc['//*[@id="formBuscaOferta:j_idt75_label"]']) 
        instruccion2[ii,1] <- try(toString(xmlValue(instruccion[[1]][[1]])))
        base[k,9] <- instruccion2[ii,1]
        ii <- ii+1
        Sys.sleep(runif (1,0.1,0.2))
        
        remuneracion <- try(doc['//*[@id="formBuscaOferta:j_idt78_label"]']) 
        remuneracion2[jj,1] <- try(toString(xmlValue(remuneracion[[1]][[1]])))
        base[k,10] <- remuneracion2[jj,1] 
        jj <- jj+1
        Sys.sleep(runif (1,0.1,0.2))
        
        experiencia <- try(doc['//*[@id="formBuscaOferta:j_idt84_label"]']) 
        experiencia2[kk,1] <- try(toString(xmlValue(experiencia[[1]][[1]])))
        base[k,11] <- experiencia2[kk,1]
        kk <- kk+1
        Sys.sleep(runif (1,0.1,0.2))
        
        area <- try(doc['//*[@id="formBuscaOferta:j_idt81_label"]']) 
        area2[ll,1] <- try(toString(xmlValue(area[[1]][[1]])))
        base[k,12] <- area2[ll,1] 
        ll <- ll+1
        Sys.sleep(runif (1,0.1,0.2))
        
        conocimiento <- try(doc['//*[@id="formBuscaOferta:j_idt72"]/tbody/tr[3]/td[2]']) 
        conocimiento2[mm,1] <- try(toString(xmlValue(conocimiento[[1]][[1]])))
        base[k,13] <- conocimiento2[mm,1]
        mm <- mm+1
        Sys.sleep(runif (1,0.1,0.2))
        
        actividades <- try(doc['//*[@id="formBuscaOferta:j_idt72"]/tbody/tr[3]/td[4]']) 
        actividades2[nn,1] <- try(toString(xmlValue(actividades[[1]][[1]])))
        base[k,14] <- actividades2[nn,1]
        nn <- nn+1
        Sys.sleep(runif (1,0.1,0.2))
        
        jornada <- try(doc['//*[@id="formBuscaOferta:j_idt94_label"]']) 
        jornada2[pp,1] <- try(toString(xmlValue(jornada[[1]][[1]])))
        base[k,15] <- jornada2[pp,1]
        pp <- pp+1
        Sys.sleep(runif (1,0.1,0.2))
        
        adicional <- try(doc['//*[@id="formBuscaOferta:j_idt72"]/tbody/tr[5]/td[2]']) 
        adicional2[qq,1] <- try(toString(xmlValue(adicional[[1]][[1]])))
        base[k,16] <- adicional2[qq,1] 
        qq <- qq+1
        Sys.sleep(runif (1,0.1,0.2))
        
        vacantes <- try(doc['//*[@id="formBuscaOferta:j_idt72"]/tbody/tr[5]/td[4]']) 
        vacantes2[rr,1] <- try(toString(xmlValue(vacantes[[1]][[1]])))
        base[k,17] <- vacantes2[rr,1] 
        rr <- rr+1
        Sys.sleep(runif (1,0.1,0.2))
        
        Sys.sleep(runif(1, 1, 2))  
        remDr$goBack()
        Sys.sleep(runif(1, 1, 2))  
        k <- k +1    
      }else{}
    }
    
  },error = function(e){})
}

save(base, file = "base.rdata")
base <- data.frame(base)
colnames(base) <- toupper(variables[c(1:dim(base)[2])])

for (i in c(1,11:16)){
  base[[i]] <- gsub("ÃƒÂ±", "ñ", base[[i]])
  base[[i]] <- gsub("ÃƒÂ³N", "ÓN", base[[i]])
  base[[i]] <- gsub("ÃƒÂ“N", "ÓN", base[[i]])
  base[[i]] <- gsub("PÃƒÂ“", "PÓ", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘O", "ÑO", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘A", "ÑA", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘A", "ÑA", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘E", "ÑE", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘I", "ÑI", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‘ÃƒÂRIA", "ÑÍA", base[[i]])
  base[[i]] <- gsub("/TÃƒÂ©", " Té", base[[i]])
  base[[i]] <- gsub("ÃƒÂ‰", "É", base[[i]])
  base[[i]] <- gsub("PEÃƒÂš", "PÉU", base[[i]])
  base[[i]] <- gsub("mÃƒÂ?a", "mía", base[[i]])
  base[[i]] <- gsub("iÃƒÂ²n", "ión", base[[i]])
  base[[i]] <- gsub("IÃƒÂ’N", "IÓN", base[[i]])
  base[[i]] <- gsub("ÃƒÂ“", "Ó", base[[i]])
  base[[i]] <- gsub("nÃƒÂ³", "nó", base[[i]])  
  base[[i]] <- gsub("ÃƒÂ³n", "ón", base[[i]])
  base[[i]] <- gsub("GÃƒÂš", "GÚ", base[[i]])  
  base[[i]] <- gsub("SÃƒÂ“", "SÓ", base[[i]])    
  base[[i]] <- gsub("GÃƒÂRIA", "GÍA", base[[i]])
  base[[i]] <- gsub("AsesorÃƒÂ?a", "Asesoría", base[[i]])
  base[[i]] <- gsub("AsesorÃƒÂa", "Asesoría", base[[i]])
  
  base[[i]] <- gsub("TÃƒÂRITULO", "TÍTULO", base[[i]])
  base[[i]] <- gsub("CÃƒÂ“", "CÓ", base[[i]])
  
  base[[i]] <- gsub("ÃƒÂA", "ÍA", base[[i]])
  base[[i]] <- gsub("ÃƒÂa", "ía", base[[i]])
  base[[i]] <- gsub("ÃƒÂ�a", "ía", base[[i]])
  
  base[[i]] <- gsub("rÃƒÂ�a", "ía", base[[i]])
  base[[i]] <- gsub("rÃƒÂa", "ría", base[[i]])
  
  
  base[[i]] <- gsub("RÃƒÂ", "RÍ", base[[i]])
  base[[i]] <- gsub("ÃƒÂ“D", "ÍD", base[[i]])
  
  base[[i]] <- gsub("ÃƒÂRIC", "ÍC", base[[i]])  
  base[[i]] <- gsub("LÃƒÂRI", "LÍ", base[[i]]) 
  base[[i]] <- gsub("TÃƒÂRI", "TÍ", base[[i]])  
  
  base[[i]] <- gsub("ÃƒÂHOP", "Á", base[[i]])   
  base[[i]] <- gsub("mÃƒÂ¡", "má", base[[i]])   
  
  base[[i]] <- gsub(" ÃƒÂ", " á", base[[i]])  
  base[[i]] <- gsub(" ÃƒÂ€", " Á", base[[i]])
  base[[i]] <- gsub("ÃƒÂRIST", "IST", base[[i]])
  
  base[[i]] <- gsub("Ã¢Â€Â¢", "-", base[[i]])
  base[[i]] <- gsub("ÃƒÂœ", "Ü", base[[i]])
  base[[i]] <- gsub("BÃƒÂ¡", "Bá", base[[i]]) 
  
  base[[i]] <- gsub("ÃƒÂ‰S", "ÉS", base[[i]])
  base[[i]] <- gsub("RÃƒÂ‰", "RÉ", base[[i]])
  base[[i]] <- gsub('"Ã‚Â°', '-', base[[i]])  
  
  base[[i]] <- gsub("PÃƒÂš", "PÚ", base[[i]])
  base[[i]] <- gsub("ÃƒÂs", "Ás", base[[i]])
  base[[i]] <- gsub("ÃƒÂS", "ÁS", base[[i]])
  base[[i]] <- gsub("IÁ“N", "IÓN", base[[i]])
  base[[i]] <- gsub("Áa", "ía", base[[i]])
  base[[i]] <- gsub("ÃƒÂ?", "í", base[[i]])
}

descarga_ex <- function(datos, file){        
  wb <- createWorkbook(type="xlsx")
  
  # Define some cell styles
  # Title and sub title styles
  TITLE_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=16, isBold=TRUE)
  
  SUB_TITLE_STYLE <- CellStyle(wb) + Font(wb,  heightInPoints=12,
                                          isItalic=TRUE, isBold=FALSE)
  
  # Styles for the data table row/column names
  TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE)
  
  TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) +
    Alignment(vertical="VERTICAL_CENTER",wrapText=TRUE, horizontal="ALIGN_CENTER") +
    Border(color="black", position=c("TOP", "BOTTOM"), 
           pen=c("BORDER_THICK", "BORDER_THICK"))+Fill(foregroundColor = "lightblue", pattern = "SOLID_FOREGROUND")
  
  sheet <- createSheet(wb, sheetName = "SOCIOEMPLEO")
  
  # Helper function to add titles
  xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
    rows <- createRow(sheet, rowIndex=rowIndex)
    sheetTitle <- createCell(rows, colIndex=1)
    setCellValue(sheetTitle[[1,1]], title)
    setCellStyle(sheetTitle[[1,1]], titleStyle)
  }
  
  # Add title and sub title into a worksheet
  xlsx.addTitle(sheet, rowIndex=4, 
                title=paste("Fecha:", format(Sys.Date(), format="%Y/%m/%d")),
                titleStyle = SUB_TITLE_STYLE)
  
  xlsx.addTitle(sheet, rowIndex=5, 
                title="Elaborado por: ",
                titleStyle = SUB_TITLE_STYLE)
  
  # Add title
  xlsx.addTitle(sheet, rowIndex=7, 
                  paste("SOCIOEMPLEO CORTE -", Sys.Date()),
                  titleStyle = TITLE_STYLE)
    
  # Add a table into a worksheet
  addDataFrame(datos,
               sheet, startRow=9, startColumn=1,
               colnamesStyle = TABLE_COLNAMES_STYLE,
               rownamesStyle = TABLE_ROWNAMES_STYLE,
               row.names = FALSE)
  
  
  # Change column width
  setColumnWidth(sheet, colIndex=c(1:ncol(datos)), colWidth=20)
  

  # Save the workbook to a file
  saveWorkbook(wb, file)
}

descarga_ex(base, file=paste0("SOCIOEMPLEO-", Sys.Date(), ".xlsx"))


