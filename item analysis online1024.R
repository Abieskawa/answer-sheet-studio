#install.packages("shiny")
#install.packages("openxlsx")
#install.packages("knitr")
#install.packages("DT")


library(shiny)
library(openxlsx)
library(knitr)
library(DT)

ui <- fluidPage(
  
  sidebarPanel(downloadButton("writedowntemp", "Download templete"),
              fileInput("file1", "uploead templete", accept = ".xlsx"),
               numericInput("point", "point of each question", 5, min = 1, max = 100),
               downloadButton("writedown", "Download correcting"),
               downloadButton("writedown4", "Download showwrong"),
               downloadButton("writedown2", "Download itemanalysis"),
               downloadButton("writedown3", "Download summary")
  ),
  mainPanel(
    tabsetPanel(
      tabPanel("Correcting",uiOutput("explanationcor"), div(DTOutput('rightornot'),style = "font-size:120%")),
      tabPanel("show wrong", div(DTOutput('showwrong'),style = "font-size:120%")),
      tabPanel("Summary",div(tableOutput("summaryTable"),style = "font-size:120%"),plotOutput("histlegend", height="250px"),plotOutput("scoreHistogram",width = "600px")),
      tabPanel("Item Analysis", div(tableOutput("itemTable"),style = "font-size:120%"))
    )
  )
)

server <- function(input, output) {
  #create temp
  number<-c('1','2','3','4')
  colone<-c('a','b','c','d')
  coltwo<-c('c','a','b','d')
  colthree<-c('c','a','b','d')
  colfour<-c('a','b','b','d')
  colfive<-c('d','c','b','a')
  templete<-data.frame(number,colone,coltwo,colthree,colfour,colfive)
  colnames(templete)<-c('number','correctanswer_write_down_below','example_student1','add_student_to_right','David','Andy')
  
  output$writedowntemp <- downloadHandler(filename = "templete.xlsx", content = function(file) {
    write.xlsx(templete, file, rowNames = FALSE)
  })
  
  output$itemTable <- renderTable({
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    
    
    orderpoint<-t(point)[order(t(point)[,2]),]
    
    studentnumber<-length(test[1,])-1
    half<-as.numeric(studentnumber%/%2)
    low<-orderpoint[1:half,]
    high<-orderpoint[-(1:half),]
    
    colnames(test[,-1])
    number<-as.numeric()
    for (k in 1:length(low[,1])) {
      number[k]<-which(colnames(test)==low[k,1])
    }
    
    
    highnumber<-as.numeric()
    for (k in 1:length(high[,1])) {
      highnumber[k]<-which(colnames(test)==high[k,1])
    }
    
    
    
    
    
    lowori<-test[,number]
    highori<-test[,highnumber]
    Difficulty<-as.numeric()
    for (i in 1:questnumber) {
      Difficulty[i]<-round(sum(test[i,-1]==test[i,1])/studentnumber,digits = 2)
    }
    
    
    Difficultyhigh<-as.numeric()
    for (i in 1:questnumber) {
      Difficultyhigh[i]<-sum(highori[i,]==test[i,1])/length(highnumber)
    }
    
    Difficultylow<-as.numeric()
    for (i in 1:questnumber) {
      Difficultylow[i]<-sum(lowori[i,]==test[i,1])/length(number)
    }
    
    
    Discrimination<-round(Difficultyhigh-Difficultylow,digits = 2)
    itemTable<-rbind(Difficulty,Discrimination)
    colnames(itemTable)<-1:questnumber
    rownames(itemTable)<-c('Difficulty','Discrimination')
    itemTable<-t(itemTable)
    
    studentanswer<-test[,-1]
    a<-as.numeric()
    b<-as.numeric()
    c<-as.numeric()
    d<-as.numeric()
    for (z in 1:questnumber) {
      a[z]<-round(sum(studentanswer[z,]=='a')/studentnumber,digits = 2)
      b[z]<-round(sum(studentanswer[z,]=='b')/studentnumber,digits = 2)
      c[z]<-round(sum(studentanswer[z,]=='c')/studentnumber,digits = 2)
      d[z]<-round(sum(studentanswer[z,]=='d')/studentnumber,digits = 2)
    }
    itemTable<-cbind(itemTable,correct,a,b,c,d)
    itemTable
  }
  , rownames = TRUE)
  

  
  output$summaryTable<-renderTable({
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)     
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    studentnumber<-length(test[1,])-1
    mean_point<-round(mean(as.numeric(point[2,])),digits = 2)
    standard_dev<-round(sd(as.numeric(point[2,])),digits = 2)
    percent88<-round(quantile(as.numeric(point[2,]),0.88, na.rm = T),digits = 2)
    percent75<-round(quantile(as.numeric(point[2,]),0.75, na.rm = T),digits = 2)
    median_point<-round(quantile(as.numeric(point[2,]),0.5, na.rm = T),digits = 2)
    percent25<-round(quantile(as.numeric(point[2,]),0.25, na.rm = T),digits = 2)
    percent12<-round(quantile(as.numeric(point[2,]),0.12, na.rm = T),digits = 2)
    
    summaryTable<-rbind(mean_point,standard_dev,percent88,percent75,median_point,percent25,percent12)
    t(summaryTable)
  }, rownames = FALSE, colnames = TRUE)
  
  output$rightornot<-renderDT({
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)     
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    color_text <- function(value1, value2) {
      if (value1 == value2) {
        return(paste0('<span style="color:black">', value1, '</span>'))
      } else {
        return(paste0('<span style="color:red">', value1, '</span>'))
      }
    }
    studentanswer<-test[,-1]
    coloranswer<-studentanswer
    for (k in 1:length(studentanswer[1,])) {
      coloranswer[,k]<- mapply(color_text, studentanswer[,k], correct)
    }
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    totalpoint<-point[2,]
    finaltable<-rbind(coloranswer,totalpoint)
    rownames(finaltable)<-c(1:questnumber,'total point')
    datatable(finaltable, escape = FALSE, options = list(pageLength = 20, searching = FALSE))
  }
  )
  
  
  output$showwrong<-renderDT({
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)     
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    showwrong<-testori
    for (k in 1:(length(testori[1,])-2)) {
      tempwrite<-vector()
      for (r in 1:length(showwrong[,2])) {
        if(testori[,2+k][r]==correct[r]){
          tempwrite[r]<-NA
        }else{
          tempwrite[r]<-showwrong[,2+k][r]
        }
      }
      showwrong[,2+k]<-tempwrite
    }
    
    datatable(showwrong, escape = FALSE, options = list(pageLength = 20, searching = FALSE))
  }
  )
  
  
  output$explanationcor<-renderUI({
    req(input$file1)
    texthere<-'Red texts mean the answers are wrong, and the red texts are the wrong answers written by the students.'
    HTML(paste0('<div style="font-size: 25px;">', texthere, '</div>'))
    
  })
  
  output$histlegend<-renderPlot({
    req(input$file1)
    plot.new()
    par(mar=c(0,0,0,0))
    legend('topleft',legend = c('mean','','median','','percentage'),lwd = 4,cex = 1.5,col = c('red','white','blue','white','black'),bty = 'n')
  })
  
  output$scoreHistogram<-renderPlot({
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)     
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    par(mar=c(5,5,0,1))
    hist(as.numeric(point[2,]),main = '',xlab = 'Point',cex.lab=2,cex.axis=2)
    abline(v=mean(as.numeric(point[2,])),lwd=2,col='red')
    abline(v=quantile(as.numeric(point[2,]),0.88, na.rm = T),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.75, na.rm = T),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.5, na.rm = T),lwd=2,col='blue')
    abline(v=quantile(as.numeric(point[2,]),0.25, na.rm = T),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.12, na.rm = T),lwd=2)
  })
  

#build download----------------------------------------------
  
  observeEvent(list(input$file1,input$point),{
    req(input$file1)
    testori <- read.xlsx(input$file1$datapath, sheet = 1)
    test <- testori[,-1]
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    
    orderpoint<-t(point)[order(t(point)[,2]),]
    
    studentnumber<-length(test[1,])-1
    half<-as.numeric(studentnumber%/%2)
    low<-orderpoint[1:half,]
    high<-orderpoint[-(1:half),]
    
    colnames(test[,-1])
    number<-as.numeric()
    for (k in 1:length(low[,1])) {
      number[k]<-which(colnames(test)==low[k,1])
    }
    
    
    highnumber<-as.numeric()
    for (k in 1:length(high[,1])) {
      highnumber[k]<-which(colnames(test)==high[k,1])
    }
    
    
    
    
    
    lowori<-test[,number]
    highori<-test[,highnumber]
    Difficulty<-as.numeric()
    for (i in 1:questnumber) {
      Difficulty[i]<-round(sum(test[i,-1]==test[i,1])/studentnumber,digits = 2)
    }
    
    
    Difficultyhigh<-as.numeric()
    for (i in 1:questnumber) {
      Difficultyhigh[i]<-sum(highori[i,]==test[i,1])/length(highnumber)
    }
    
    Difficultylow<-as.numeric()
    for (i in 1:questnumber) {
      Difficultylow[i]<-sum(lowori[i,]==test[i,1])/length(number)
    }
    
    
    Discrimination<-round(Difficultyhigh-Difficultylow,digits = 2)
    itemTable<-rbind(Difficulty,Discrimination)
    colnames(itemTable)<-1:questnumber
    rownames(itemTable)<-c('Difficulty','Discrimination')
    itemTable<-t(itemTable)
    
    studentanswer<-test[,-1]
    a<-as.numeric()
    b<-as.numeric()
    c<-as.numeric()
    d<-as.numeric()
    for (z in 1:questnumber) {
      a[z]<-round(sum(studentanswer[z,]=='a')/studentnumber,digits = 2)
      b[z]<-round(sum(studentanswer[z,]=='b')/studentnumber,digits = 2)
      c[z]<-round(sum(studentanswer[z,]=='c')/studentnumber,digits = 2)
      d[z]<-round(sum(studentanswer[z,]=='d')/studentnumber,digits = 2)
    }
    itemTable<-cbind(itemTable,correct,a,b,c,d)
    
    studentnumber<-length(test[1,])-1
    mean_point<-round(mean(as.numeric(point[2,])),digits = 2)
    standard_dev<-round(sd(as.numeric(point[2,])),digits = 2)
    percent88<-round(quantile(as.numeric(point[2,]),0.88, na.rm = T),digits = 2)
    percent75<-round(quantile(as.numeric(point[2,]),0.75, na.rm = T),digits = 2)
    median_point<-round(quantile(as.numeric(point[2,]),0.5, na.rm = T),digits = 2)
    percent25<-round(quantile(as.numeric(point[2,]),0.25, na.rm = T),digits = 2)
    percent12<-round(quantile(as.numeric(point[2,]),0.12, na.rm = T),digits = 2)
    
    summaryTable<-rbind(mean_point,standard_dev,percent88,percent75,median_point,percent25,percent12)
    
    totalpoint<-point[2,]
    resulttablepre<-rbind(studentanswer,totalpoint)
    correctanswer<-c(correct,NA)
    resulttable<-cbind(correctanswer,resulttablepre)
    rownames(resulttable)<-c(1:questnumber,'total point')
    
    showwrong<-testori
    for (k in 1:(length(testori[1,])-2)) {
      tempwrite<-vector()
      for (r in 1:length(showwrong[,2])) {
        if(testori[,2+k][r]==correct[r]){
          tempwrite[r]<-NA
        }else{
          tempwrite[r]<-showwrong[,2+k][r]
        }
      }
      showwrong[,2+k]<-tempwrite
    }
    
    #downloadhandler----------------------------------------------------
    
    output$writedown <- downloadHandler(filename = "Correcting.xlsx", content = function(file) {
      write.xlsx(resulttable, file, rowNames = T)
    })
    output$writedown2 <- downloadHandler(filename = "Item Analysis.xlsx", content = function(file) {
      write.xlsx(itemTable, file, rowNames = T)
    })
    output$writedown3 <- downloadHandler(filename = "Summary.xlsx", content = function(file) {
      write.xlsx(t(summaryTable), file, rowNames = F)
    })
    output$writedown4 <- downloadHandler(filename = "showwrong.xlsx", content = function(file) {
      write.xlsx(showwrong, file, rowNames = F)
    })
    
    
  })
}

# Run the application 
shinyApp(ui = ui, server = server)
