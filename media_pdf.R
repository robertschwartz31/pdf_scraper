### All relevant packages associated with this
### app. I'm actually not sure if I use stringr anywhere
### here, but it's never a bad tool to have
library(shiny)
library(pdftools)
library(stringr)
library(dplyr)
library(shinyFiles)
library(shinyWidgets)
library(shinydashboard)
library(openxlsx)
library(formattable)

### This script isn't as messy as Brad's Selenium script, but it's also nowhere near as clean as scrubbing
### DialogTech, LeadExec, or Google data. It's a farcry from bad or ineffective, it just has its limits.
### Its main limitation is that this doesn't particularly work well with non culligan clients, in terms 
### of the pdf scraping anyway. The actual formatting and billing collection is pretty nice, it's just the
### pdf scraping of non culligan that is problematic. In any event, this does a pretty ok job at getting 
### billing stuff together, it just needs extra skilled eyes to parce more of the details 

### Just the pretty buttons that we're used to
primaryActionButton <- function(inputId, label) tags$button(id = inputId, type = "button", class = "btn btn-primary action-button btn-lg", label)
successDownloadButton <- function(outputId, label) aTag <- tags$a(id = outputId, class = "btn btn-success shiny-download-link btn-lg", href = "", target = "_blank", icon("download"), label)

### This needs to be done so that old data isn't reused in next iterations
pdf_data_frame <<- NULL

### Dashboards just look nicer overall, and they're easy to build
ui <- dashboardPage(
  
  ### Title. Just a working Title
  dashboardHeader(title = 'Media Files'), 
  
  ### Builds the sidebar items. There's only 2 for right now.
  dashboardSidebar(
    
    sidebarMenu(id = 'tabs', 
                
                menuItem('PDF RenameR', tabName = 'renameR', icon = icon('file-pdf')),
                menuItem('SpreadSheet mergeR', tabName = 'mergeR', icon = icon('file-export'))
                
                )
    
  ), 
  
  ### Here's where the UI for features is actually built
  dashboardBody(
    
    ### Just a container for all (2) items
    tabItems(
      
      ### Builds the UI for the PDF renaming tool
      tabItem(tabName = 'renameR', 
              
                ### Header, minor description, button, and directory select
                h2('Get PDFs'),
                helpText("The PDF's can be grabbed in bulk. The goal of this will be to edit and extract pdf files in a folder"),
                primaryActionButton(inputId = 'format_pdfs', label = 'Format PDFs'),
                shinyDirButton("folder", "Change Folder", "Please select a folder")
              
              ),
      
      ### Builds the UI for the File merging tool
      tabItem(tabName = 'mergeR', 
              
              ### Header, minor description, csv bin, and button
              h2('Get Spreadsheet'), 
              helpText('Make sure that the spreadsheet input is a csv. It makes the data much easier to read.'),
              fileInput(inputId = 'media_raw', label = 'Input Media csv:'), 
              primaryActionButton(inputId = 'get_media_xlsx', label = 'Get Media Spreadsheet'),
              br(),
              uiOutput('finish_merge')
               
      )
      
    )
    
  )
  
)

### Where the work actually gets done
server <- function(input, output, session) {
  
  ### File paths in Mac and Windows are handled differently, 
  ### so this attempts to get the correct pathway
  sysinfo <- Sys.info()
  is_windows <- switch(tolower(sysinfo["sysname"]), windows = T, F)
  
  ### Windows is absolute trash. I can't believe how terrible it is at not having generalized shortcuts,
  ### no wonder why devs pretty much only use UNIX
  path_to_desktop <- ifelse(is_windows, paste0(str_extract(getwd(), '^C\\:\\/.*\\/.*\\/'), 'Desktop'), '~/Desktop')
  
  ### There's a lot that rides on this function. It gets the proper naming convention from the PDF, and it edits the 
  ### clients in the raw spreadsheet for better code merging. If more clients need to be added, this is where it will be
  get_client_name <- function(account_string) {
    
    ### we want to be able to figure out which client the thing applies to
    client_list <- c('Marling', 'Culligan', 'Cost Cutters', 'Anytime Fitness', 'WI Book Festival',
                     'Habitat for Humanity', 'Habitat ReStore', 'Gingras|Thomsen|Wach|GTW Lawyers', 'Sergenians', 
                     'Event Essentials', 'Wisconsin Book Festival', 'Madison Public Library', 'Phantom Fireworks'
                     )
    
    ### Finds which client is in the pdf
    selected_client <- client_list[str_detect(str_to_lower(account_string), str_to_lower(client_list))]
    
    if(length(selected_client) == 0) browser() # return('Unknown')
    
    ### Makes sure to rename Lawyer stuff to GTW lawyers
    if(selected_client == 'Gingras|Thomsen|Wach|GTW Lawyers') selected_client <- 'GTW Lawyers'
    if(selected_client == 'Wisconsin Book Festival' | selected_client ==  'Madison Public Library') selected_client <- 'WI Book Festival'
    
    ### Returns the correct client
    selected_client
    
  }
  
  ### Handles when the shiny directory button is selected
  observeEvent(input$folder, {
    
    ### OK, so this is pretty easy, just select the folder you want. We'll need to have a discussion about what path the files
    ### can be saved to
    if(dir.exists('/Volumes/Front/Adam/Reporting/')) {
      file_paths <- c('Desktop' = path_to_desktop, 'Volumes' = '/Volumes/Front/Adam/Reporting/')
    }else file_paths <- c('Desktop' = path_to_desktop)
    
    ### Actually allows the file path to be selected
    shinyDirChoose(input, "folder", roots = file_paths)
    
  })
  
  ### Reads and formats the pdf's
  observeEvent(input$format_pdfs, {
    
    ### Returns warning if there is no directory selected
    if(class(input$folder) != 'list') return(sendSweetAlert(session = session, 
                                                            title = 'Please choose the pathway to the desired pdf folder!',
                                                            type = 'warning'))
    
    ### Progress indicators are always good
    withProgress(message = "Renaming PDF's", value = 0, {
   
    ### This is likely to change, but it makes the most sense on my directory. I'll play around on my laptop for later testing
    selected_path <- paste0(path_to_desktop, paste0(input$folder$path, collapse = "/"), '/')
    
    ### Since I'm renaming files, I don't want to skip over anything
    original_list_files_main <- list.files(selected_path)
    
    ### Returns a warning if the directory does not only contain pdf files
    if(!all(grepl('.*\\.pdf', original_list_files_main))) return(sendSweetAlert(session = session, 
                                                                           title = 'Make sure that the file only contains pdf files!',
                                                                           text = 'Having other file types in this folder will make this not work',
                                                                           type = 'warning'))
    
    ### Gets the length vector
    length_vector <- 1:length(list.files(path = selected_path))
    n_pdf <- length(list.files(path = selected_path)) + 3
    
    ### Initial file rename
    file.rename(paste0(selected_path, list.files(path = selected_path, pattern = '.*\\.pdf')), paste0(selected_path, length_vector, gsub('\\-.*[[:digit:]].*[[:digit:]]', '', list.files(path = selected_path))))
    
    ### Since I'm renaming files, I don't want to skip over anything. BAD NAMING CONVENTION
    original_list_files <- list.files(selected_path)
    
    ### Probably won't even be seen
    incProgress(1/n_pdf, message = "Renaming Masses")
    
    ### Creates an empty data frame for things to be scrubbed to
    pdf_data_frame <- data.frame()
    
    ### This is clearly stolen from selenium_script.R because there are some Windows vs Mac distinctions. Windows sucks.
    line_break <- ifelse(is_windows, "\\\r\\\n", "\\\n")
    
    ### Iterates across each file
    break_loop <- F
    
    for (i in length_vector) {
  
      ### Default positions for 0 cost and non culligan
      zero_net <- F
      non_culligan <- F
      
      ### Reads the pdf as a single string. Makes sure subsequent pages are included
      new_pdf <- paste0(pdf_text(paste0(selected_path, original_list_files[i])), collapse = '')
      
      ### Breaks the pdf into a vector by line
      new_pdf <- gsub("^ *", "", strsplit(new_pdf, line_break)[[1]])
      
      ### Forces space items to a single space
      new_pdf <- gsub("\\s{2,}", " ", new_pdf)
      
      ### Finds value of Net Due to see if a $0 tag should be appended
      net_due <- gsub('^Net\\sDue\\:\\s', '', new_pdf[grepl('Net\\sDue\\:\\s\\$', new_pdf)])[1]
      if(net_due == '$0.00') zero_net <- T
      
      ### Gets the line providing 
      account <- new_pdf[grep('Account Executive Advertiser Product Order Type', new_pdf) + 1]
      
      ### Gets the client (I don't think the tryCatch() is necessary anymore)
      pdf_client <- tryCatch(get_client_name(account))
      
      ### Breaks the loop and notifies that the loop has been broken if the client is not in the 
      ### predefined list
      if(pdf_client == 'Unknown') return({
        
        break_loop <- T
        break
        
      })
      
      ### This isn't technically the first description line, but it's how these positions are found
      first_description <- grep('description', new_pdf, ignore.case = T)
      
      ### Sometimes the description isn't shown if the file is short. Only fires if the description is found
      if(length(first_description) > 0) {
        
        first_description <- min(grep('description', new_pdf, ignore.case = T)) - 2
        format_new_pdf <- new_pdf[first_description:length(new_pdf)]
        
      }else format_new_pdf <- new_pdf
      
      ### Finds and removes additional comments section, if it exists
      additional_comments <- max(grep('additional\\scomments\\:', format_new_pdf, ignore.case = T))
      if(length(additional_comments) != 0) format_new_pdf <- format_new_pdf[1:(additional_comments - 1)]
      
      ### Gets the list of lines not relevant to the data
      junk_lines <- grep('Line Start End Days Spots\\/Week Rate', format_new_pdf)
      page_label <- grep('Page\\s.*\\sof\\s', format_new_pdf)
      headers <- grep("\\# Day Date Time Length Rate Copy Program Description Class Remarks", format_new_pdf)
      description_tag <- grep('Description\\:', format_new_pdf)
      junk_lines <- c(junk_lines, junk_lines + 1, page_label, headers, description_tag)
      junk_lines <- junk_lines[order(junk_lines)]
      
      ### Removes the junk lines
      format_new_pdf <- format_new_pdf[-junk_lines]
      
      ### If the comments section still exists, it means scrubbing failed
      if(any(format_new_pdf == 'Comments:')) {
        
        ### Returns not found if the comments section still exists
        script_codes <- 'Not Found'
        
      }else{
      
        ### Splits for better formatting
        break_string <- strsplit(format_new_pdf, '\\s')
        
        ### Iterates across the list for the unique script codes
        for (index in 1:length(break_string)) {
          
          ### Chooses the 8th item (should be the corresponding code)
          break_string[[index]] <- break_string[[index]][8]
          
        }
        
        ### Gets the unique script codes
        script_codes <- unique(unlist(break_string))
        script_codes <- script_codes[!is.na(script_codes)]
        script_codes <- script_codes[order(script_codes)]
        
      }
      
      ### Grabs the station name
      station_name <- gsub('\\-.*$', '', new_pdf[1])
      
      ### Boolean for whether it is culligan or not
      is_culligan <- ifelse(pdf_client == 'Culligan', T, F)
      
      ### Grabs the new entry frame
      new_entry <- data.frame('Station' = station_name, 'script_code' = script_codes, 'is_culligan' = is_culligan, 'Client' = pdf_client, stringsAsFactors = F)
      pdf_data_frame <- rbind(pdf_data_frame, new_entry)
      
      ### Changes the non Culligan if it's not culligan
      if(pdf_client != 'Culligan') non_culligan <- T
      
      ### These rename differently if there is a zero net cost or if it is a non culligan client
      if(non_culligan | zero_net) {
        
        if(non_culligan & zero_net) {
          
          ### Naming if 0 and non culligan
          file.rename(paste0(selected_path, original_list_files[i]), paste0(selected_path, gsub('\\.pdf', '', original_list_files[i]), ' - ', pdf_client, ' - $0.pdf'))
          
        }else{
          
          ### Naming if nonculligan
          if(non_culligan) file.rename(paste0(selected_path, original_list_files[i]), paste0(selected_path, gsub('\\.pdf', '', original_list_files[i]), ' - ', pdf_client, '.pdf'))
          
          ### Naming if 0 cost
          if(zero_net) file.rename(paste0(selected_path, original_list_files[i]), paste0(selected_path, gsub('\\.pdf', '', original_list_files[i]), ' - $0.pdf'))
          
        }
        
      }
      
      ### Just to show which iteration
      incProgress(1/n_pdf, message = paste0('Scraping PDF ', i, '/', n_pdf - 3))
      
    }
    
    ### Probably won't be seen
    incProgress(1/n_pdf, message = "Indexing Copies")
    
    ### Gets the new file names
    new_file_names <- gsub('^[0-9]*|\\.pdf$', '', list.files(path = selected_path))
    
    ### Gets the duplicated and unique station names
    duplicated_stations <- new_file_names[duplicated(new_file_names)]
    unique_duplicated_stations <- unique(new_file_names[duplicated(new_file_names)])
    
    ### Creates a new frame showing station and file path
    ### pattern = paste0(duplicated_stations, collapse = '|')
    all_dup_file_paths <- data.frame('Files' = paste0(selected_path, list.files(path = selected_path)), stringsAsFactors = F)
    all_dup_file_paths$new_file_names <- new_file_names
    all_dup_file_paths$Station <- gsub('^.*\\/[0-9]*|\\.pdf$', '', all_dup_file_paths$Files)
    all_dup_file_paths <- all_dup_file_paths[all_dup_file_paths$new_file_names %in% duplicated_stations, ]
    
    ### Creates a new naming convention. I could probably make this better, but it still denotes things properly enough that I don't expect any problems to be had.
    all_dup_file_paths <- all_dup_file_paths %>% 
      group_by(Station) %>% 
      mutate('Duplicate' = row_number() - 1) %>% 
      filter(Duplicate != 0) %>% 
      mutate('new_name' = paste0(gsub('\\.pdf$', '', Files), ' (', Duplicate, ').pdf'))
    
    ### Renames the new files to give them numeric indexes
    file.rename(all_dup_file_paths$Files, all_dup_file_paths$new_name)  
    
    ### Strips files of their leading numbers
    incProgress(1/n_pdf, message = "Removing Dummy Numbers")
    file.rename(paste0(selected_path, list.files(path = selected_path, pattern = '.*\\.pdf')), paste0(selected_path, gsub('^[0-9]*', '', list.files(path = selected_path, pattern = '.*\\.pdf'))))
    
    })
    
    ### I want this to be accessible later down the funnel
    pdf_data_frame <<- pdf_data_frame

    ### Here's where we start indexing stations with multiple entries
    
    if(break_loop) { 
      
      file.rename(paste0(selected_path, list.files(selected_path)), paste0(selected_path, original_list_files_main))
      
      sendSweetAlert(session = session, 
                     title = 'Unknown Client Name', 
                     text = 'Please alert the Data Team to update the client list!',
                     type = 'error')
      
    }else sendSweetAlert(session = session, 
                         title = 'PDF Scraping and Renaming Complete!',
                         type = 'success')
    
   
  })
  
  ### Creates the Billing xlsx file
  observeEvent(input$get_media_xlsx, {
    
    ### I need the pdf_data_frame to be available for anything productive to actually happen
    if(is.null(pdf_data_frame)) return(sendSweetAlert(session = session, 
                                                      title = 'Table from PDF scrape not found!', 
                                                      text = 'Make sure that the RenameR feature is used first', 
                                                      type = 'warning'))
    
    ### Returns a warning if there is no input file
    if(is.null(input$media_raw)) return(sendSweetAlert(session = session, 
                                                       title = 'No media file has been input!', 
                                                       type = 'warning'))
    
    ### Returns a warning if the file is not a csv file
    if(!grepl('\\.csv$', input$media_raw$datapath)) return(sendSweetAlert(session = session, 
                                                       title = 'This is not a csv file!', 
                                                       text = 'Make sure that the RenameR feature is used first',
                                                       type = 'warning'))
    
    ### Reads the file, but wraps it in a tryCatch() in case there's something wrong with the file input
    raw_media <- tryCatch(read.csv(input$media_raw$datapath, stringsAsFactors = F))
    
    ### Returns a warning if the csv file can't be read
    if(class(raw_media) == 'try-error') return(sendSweetAlert(session = session, 
                                                              title = 'There was an error reading the csv file!', 
                                                              text = 'Try again, and if the problem persists, contact the Data Team',
                                                              type = 'warning'))
    
    ### These are the necessary columns needed for the doc. I don't expect this to change, but it's easy to adjust if it does                                                                    
    media_column_names <- c('Market', 'Client', 'Station', 'Inv.', 'InvGross', 'InvNet', 'SchGross', 'SchNet')
    
    ### Returns a warning if the column names 
    if(!all(media_column_names %in% colnames(raw_media))) return(sendSweetAlert(session = session, 
                                                                                title = 'This does not seem to be the proper file!', 
                                                                                text = 'Make sure to include column names. If there are still issues, contact the Data Team',
                                                                                type = 'warning'))
    
    ### This gets the media file down to about what it needs to be 
    main_media <- raw_media %>% 
      select(Market, Client, Station, Inv., InvGross, InvNet, SchGross, SchNet) %>% 
      mutate(InvGross = as.numeric(gsub('\\$|\\,', '', InvGross))) %>% 
      mutate(InvNet = as.numeric(gsub('\\$|\\,', '', InvNet))) %>% 
      mutate(SchGross = as.numeric(gsub('\\$|\\,', '', SchGross))) %>% 
      mutate(SchNet = as.numeric(gsub('\\$|\\,', '', SchNet))) %>% 
      group_by(Market, Client, Station, Inv.) %>% 
      summarise(InvGross = sum(InvGross), InvNet = sum(InvNet), SchGross = sum(SchGross), SchNet = sum(SchNet)) %>% 
      filter(SchGross != 0 & SchNet != 0) %>%
      ungroup() %>% 
      mutate('end_label' = gsub('^.*\\-', '', Station)) %>% 
      mutate('Station' = gsub('\\-.*$', '', Station))
    
    ### Separates media buy culligan and non culligan
    culligan_media <- main_media %>% filter(grepl('culligan', Client, ignore.case = T))
    non_culligan_media <- main_media %>% filter(!grepl('culligan', Client, ignore.case = T))
    
    ### This should make sure at least that the blatantly incorrectly pulled codes are removed
    pdf_extract <- pdf_data_frame %>% 
      mutate('script_code' = gsub('\\/$', '', script_code)) %>% 
      distinct() %>% 
      filter(grepl('[[:alpha:]]+', script_code) & grepl('[[:digit:]]+', script_code)) 
    
    ### Sorry for all the duplicates here for things that are essentially identical, but I really need the 
    ### culligan and non culligan data to be separated. It's not a great system, but this is just the way 
    ### it currently has to be. Maybe one day I'll be smart and think of a better way to handle, but this is
    ### the best I can do to handle everything
    
    ### There seems to be carry over between Culligan and Non Culligan Clients, so they need to be separated
    ### This is the culligan station / codes
    culligan_pdf_extract <- pdf_extract %>% 
      filter(is_culligan == T) %>% 
      select(-is_culligan) %>% 
      group_by(Station) %>% 
      arrange(script_code) %>% 
      summarise(script_code = paste0(script_code, collapse = '/'))
    
    ### This is the non-culligan station / codes
    non_culligan_pdf_extract <- pdf_extract %>% 
      filter(is_culligan == F) %>% 
      select(-is_culligan) %>% 
      group_by(Station, Client) %>% 
      arrange(script_code) %>% 
      summarise(script_code = paste0(script_code, collapse = '/'))
    
    ### Merges the codes with the Market, Culligan only. Moves the end labels back
    culligan_media <- left_join(culligan_media, culligan_pdf_extract, by = 'Station') %>% 
      mutate(Station = paste0(Station, '-', end_label)) %>% 
      select(-end_label) %>% 
      mutate(script_code = ifelse(is.na(script_code), 'Not Found', script_code)) %>% 
      mutate(InvGross = as.numeric(InvGross)) %>% 
      mutate(InvNet = as.numeric(InvNet)) %>% 
      mutate('Discrepencies' = ifelse(SchNet - InvNet >= 0, paste0('$', comma(SchNet - InvNet, digits = 2)), paste0('-$', comma(abs(SchNet - InvNet), digits = 2)))) %>% 
      group_by(Market) %>% 
      mutate('Gross_Total' = sum(InvGross)) %>% 
      mutate('first_iteration' = as.character(row_number())) %>% 
      mutate('last_iteration' = as.character(eval(row_number() == max(row_number())))) %>%
      mutate('InvGross' = paste0('$', comma(as.numeric(InvGross), digits = 2))) %>%
      mutate('InvNet' = paste0('$', comma(as.numeric(InvNet), digits = 2))) %>%
      mutate('SchGross' = paste0('$', comma(as.numeric(SchGross), digits = 2))) %>%
      mutate('SchNet' = paste0('$', comma(as.numeric(SchNet), digits = 2))) %>%
      mutate('Gross_Total' = paste0('$', comma(as.numeric(Gross_Total), digits = 2))) %>%
      # mutate('InvGross' = as.character(digits(as.numeric(InvGross), digits = 2))) %>%
      # mutate('InvNet' = as.character(digits(as.numeric(InvNet), digits = 2))) %>%
      # mutate('SchGross' = as.character(digits(as.numeric(SchGross), digits = 2))) %>%
      # mutate('SchNet' = as.character(digits(as.numeric(SchNet), digits = 2))) %>%
      # mutate('Gross_Total' = as.character(digits(as.numeric(Gross_Total), digits = 2))) %>%
      mutate('Accounting_Notes' = '') %>% 
      mutate('Other_Notes' = '') %>% 
      mutate('Media_Notes' = '') %>% 
      ungroup() %>% 
      select(Market, Client, Station, Inv., InvGross, InvNet, SchGross, SchNet, Gross_Total,
             Accounting_Notes, Other_Notes, script_code, Discrepencies, Media_Notes, 
             first_iteration, last_iteration)
    
    ### This renames the clients in the non_culligan_media file so everything can be merged
    ### more effectively. It still isn't great, but it ensures that codes running in the same
    ### market for different clients don't merge together
    
    ### Iterates across the unique names of the clients
    for (client in unique(non_culligan_media$Client)) {
      
      ### Renames the client names to the names the pdf's were renamed to
      non_culligan_media$Client[non_culligan_media$Client == client] <- get_client_name(client)
      
    }
    
    ### Merges the codes with the Market, non-Culligan only. Moves the end labels back
    non_culligan_media <- left_join(non_culligan_media, non_culligan_pdf_extract, by = c('Station', 'Client')) %>% 
      mutate(Station = paste0(Station, '-', end_label)) %>% 
      mutate('Client' = gsub('\\s\\-.*$', '', Client)) %>% 
      select(-end_label) %>% 
      mutate(script_code = ifelse(is.na(script_code), 'Not Found', script_code)) %>% 
      mutate(InvGross = as.numeric(InvGross)) %>% 
      mutate(InvNet = as.numeric(InvNet)) %>% 
      mutate('Discrepencies' = ifelse(SchNet - InvNet >= 0, paste0('$', comma(SchNet - InvNet, digits = 2)), paste0('-$', comma(abs(SchNet - InvNet), digits = 2)))) %>% 
      group_by(Market, Client) %>% 
      mutate('Gross_Total' = sum(InvGross)) %>% 
      mutate('first_iteration' = as.character(row_number())) %>% 
      mutate('last_iteration' = as.character(eval(row_number() == max(row_number())))) %>%
      mutate('InvGross' = paste0('$', comma(as.numeric(InvGross), digits = 2))) %>%
      mutate('InvNet' = paste0('$', comma(as.numeric(InvNet), digits = 2))) %>%
      mutate('SchGross' = paste0('$', comma(as.numeric(SchGross), digits = 2))) %>%
      mutate('SchNet' = paste0('$', comma(as.numeric(SchNet), digits = 2))) %>%
      mutate('Gross_Total' = paste0('$', comma(as.numeric(Gross_Total), digits = 2))) %>%
      # mutate('InvGross' = as.character(digits(as.numeric(InvGross), digits = 2))) %>%
      # mutate('InvNet' = as.character(digits(as.numeric(InvNet), digits = 2))) %>%
      # mutate('SchGross' = as.character(digits(as.numeric(SchGross), digits = 2))) %>%
      # mutate('SchNet' = as.character(digits(as.numeric(SchNet), digits = 2))) %>%
      # mutate('Gross_Total' = as.character(digits(as.numeric(Gross_Total), digits = 2))) %>%
      mutate('Accounting_Notes' = '') %>% 
      mutate('Other_Notes' = '') %>% 
      mutate('Media_Notes' = '') %>% 
      ungroup() %>% 
      select(Market, Client, Station, Inv., InvGross, InvNet, SchGross, SchNet, Gross_Total,
             Accounting_Notes, Other_Notes, script_code, Discrepencies, Media_Notes, 
             first_iteration, last_iteration) %>% 
      arrange(Client)
    
    ### Binds culligan and non culligan
    full_media <- rbind(culligan_media, non_culligan_media)
    full_media$unique <- paste0(full_media$Market, ' ', full_media$Client)
    
    ### Creates empty frames for binding
    empty_row <- data.frame(t(data.frame(rep('', 17), stringsAsFactors = F)), stringsAsFactors = F, row.names = NULL)
    empty_frame_two <- rbind(empty_row, empty_row)
    empty_frame_three <- rbind(empty_frame_two, empty_row)
    
    ### Mocks the same column names as the media file
    colnames(empty_row) <- colnames(full_media)
    colnames(empty_frame_two) <- colnames(full_media)
    colnames(empty_frame_three) <- colnames(full_media)
    
    ### Starts restructring the frame to be more like the original
    restructured_media <- data.frame()
    unique_markets <- unique(full_media$unique)
    
    ### Iterates across unique markets and clients
    for (i in 1:length(unique_markets)) {
      
      ### Gets the previous entry, if it exists
      if(i == 1) {
        prior_entry <- empty_row
      }else prior_entry <- full_media %>% filter(unique == unique_markets[i - 1])
      
      ### Gets the new entry and binds a row, if it's not the final row
      new_entry <- full_media %>% filter(unique == unique_markets[i])
      if(i != length(unique_markets)) new_entry <- rbind(new_entry, empty_row)
      
      ### Gets the unique client for the current iteration
      new_unique <- unique(new_entry$Client)
      new_unique <- new_unique[new_unique != '']
      
      ### Gets the unique client for the prior iteration. Returns empty if there is no prior iteration
      prior_unique <- unique(prior_entry$Client)
      prior_unique <- prior_unique[prior_unique != '']
      if(length(prior_unique) == 0) prior_unique <- ''
      
      ### Binds a longer row if there's a different unique client
      if(new_unique != prior_unique & prior_unique != '') new_entry <- rbind(empty_frame_two, new_entry)
      
      ### Binds the restructured data together
      restructured_media <- rbind(restructured_media, new_entry)
      
    }
    
    ### Gets some relevant indicies for removing excess data
    not_first_instances <- grep('^1$', restructured_media$first_iteration, invert = T)
    not_last_instance <- grep('TRUE', restructured_media$last_iteration, invert = T)
    
    ### Gets the first and last instances for the table (to be used for formatting)
    first_instances <- grep('^1$', restructured_media$first_iteration)
    last_instances <- grep('TRUE', restructured_media$last_iteration)
    
    ### I assume these will need to be formatted differently, but we'll see.
    single_length <- intersect(first_instances, last_instances)
    
    ### Removes excess data
    restructured_media[not_first_instances, c(1, 2)] <- ''
    restructured_media[not_last_instance, 9] <- ''
    
    ### Removes the iteration columns since they're not necessary
    restructured_media <- restructured_media %>% 
      select(-first_iteration, -last_iteration, -unique)
    
    ### Forces the column names to be much more similar to that of the 
    colnames(restructured_media) <- c('Market', 'Client', 'Station', 'Inv#', 'InvGross', 
                                      'InvNet', 'SchGross', 'SchNet', 'Gross Total', 
                                      'Accounting Notes', 'Other Notes', 'Scripts', 'Discrepancies',
                                      'Media Notes')
    
    ### I'm afraid we're gonna have to use...MATH
    ### Gets the indicies where there are only empty values. Need to add 1 because of the column names
    empty_indices <- which(rowSums(restructured_media == '') == 14) + 1
    odd_empty_indicies <- empty_indices[empty_indices %% 2 == 1]
    even_empty_indices <- empty_indices[empty_indices %% 2 == 0]
    
    xlsx_last_instances <- last_instances + 1
    single_length_indices <- single_length + 1
    
    negative_indices <- which(grepl('\\-', restructured_media$Discrepancies)) + 1
    
    ### Function that gets an empty row between 2 other empty rows
    get_sandwiched_indices <- function(vector) {
      
      ### Creates the main vector
      new_vector <- c()
      
      ### Eh, I guess it wasn't all that mathy. Iterates across vector length
      for (i in 1:length(vector)) {
        
        ### Gets the explicit vector value
        vector_value <- vector[i]
        
        ### Checks to see if a value above or below actually exists
        if(length(vector[i - 1]) == 0) next
        if(length(vector[i + 1]) == 0) next
        
        ### Checks to see if it's actually sandwiched between two values
        if(!((vector_value - 1) %in% vector)) next
        if(!((vector_value + 1) %in% vector)) next
        
        ### Adds a new entry if it actually is sandwiched between two values
        new_vector <- c(new_vector, vector_value)
        
      }
      
      ### Returns the vector
      new_vector 
      
    }
    
    ### Gets the list of sandwiched indices
    sandwiched_indices <- get_sandwiched_indices(empty_indices)
    
    ### Function that adds a formatting to certain rows / columns of a spreadsheet
    ### (only for the summary though). I learned about stack WAAAY to late, though
    ### there are plenty of practical applications where I do want formats to be overwritten
    format_spreadsheet <- function(x, format, row, column, loop = F, stack = F) {
      
      if(!loop) {
        
        ### Adds style to one specific column
        addStyle(x, sheet = 1, format, rows = row, cols = column, stack = stack)
        
      }else{
        
        for (iter in column) {
          
          ### Adds styles to multiple columns
          addStyle(x, sheet = 1, format, rows = row, cols = iter, stack = stack)
          
        }
      }
      
      ### Returns x so the function can be pipeable
      x
      
    }
    
    ### Here lies all of the formatting options. This doesn't need to get too fancy
    ### Supposedly, I'm supposed to be able to change the classes of the columns for easy
    ### adding, but I clearly cannot figure this out. I don't really know why, but that's just how it is
    header_border_underline <- createStyle(fgFill = '#DBE0F0', textDecoration = 'bold', border = 'bottom')
    header_no_border <- createStyle(fgFill = '#DBE0F0', textDecoration = 'bold')
    currency_type <- createStyle(numFmt = "CURRENCY")
    blue_standard <- createStyle(fgFill = '#DBE0F0')
    blue_spacer <- createStyle(fgFill = '#DBE0F0', border = c('top', 'bottom'))
    white_spacer <- createStyle(border = c('top', 'bottom'))
    right_white_border <- createStyle(border = 'right')
    right_blue_border <- createStyle(border = 'right', fgFill = '#DBE0F0')
    top_border <- createStyle(border = 'top')
    grand_total_right_corner <- createStyle(fgFill = '#4DACEB', border = c('bottom', 'right'))
    grand_total_box <- createStyle(fgFill = '#4DACEB', border = c('top', 'right', 'bottom'))
    client_separator <- createStyle(fgFill = '#253661', border = c('top', 'bottom'), borderColour = '#253661')
    red_text <- createStyle(fontColour = 'red')

    ### Bunch of row items that know where to format different 
    total_rows_vector <- 1:(nrow(restructured_media) + 1)
    odd_rows <- total_rows_vector[total_rows_vector %% 2 == 1]
    even_rows <- total_rows_vector[total_rows_vector %% 2 == 0]
    not_gross_total <- c(1:8, 10:14)
    currency_columns <- c(5:9, 13)
    not_currency_columns <- c(1:4, 10:12, 14)

    ### Starts to create the workbook
    media_xlsx <- createWorkbook()
    
    ### Adds the Billing Summary to the first tab
    addWorksheet(wb = media_xlsx, sheetName = 'Billing Summary', gridLines = T)
    
    ### I thought maybe that this would preset the type, but it doesn't look like that's the case.
    ### Maybe someday I'll figure it out, but for right now I have no idea why this doesn't work
    # format_spreadsheet(media_xlsx, currency_type, row = total_rows_vector, column = currency_columns, loop = T)
    restructured_media$`Gross Total` <- as.numeric(gsub('\\$|\\,', '', restructured_media$`Gross Total`))
    restructured_media$InvGross <- as.numeric(gsub('\\$|\\,', '', restructured_media$InvGross))
    restructured_media$InvNet <- as.numeric(gsub('\\$|\\,', '', restructured_media$InvNet))
    restructured_media$SchGross <- as.numeric(gsub('\\$|\\,', '', restructured_media$SchGross))
    restructured_media$SchNet <- as.numeric(gsub('\\$|\\,', '', restructured_media$SchNet))
    restructured_media$Discrepancies <- as.numeric(gsub('\\$|\\,', '', restructured_media$Discrepancies))

    writeData(wb = media_xlsx, sheet = 1, x = restructured_media)
    
    ### Saves the format settings to the document
    format_spreadsheet(media_xlsx, blue_standard, row = odd_rows, column = not_gross_total, loop = T) %>% 
      format_spreadsheet(right_blue_border, row = odd_rows, column = 9) %>%
      format_spreadsheet(right_white_border, row = even_rows, column = 9) %>% 
      format_spreadsheet(format = header_border_underline, row = 1, column = 1:9, loop = T) %>% 
      format_spreadsheet(format = header_no_border, row = 1, column = 10:14, loop = T) %>% 
      format_spreadsheet(top_border, row = nrow(restructured_media) + 2, column = 1:9, loop = T) %>% 
      format_spreadsheet(blue_spacer, row = odd_empty_indicies, column = 1:9, loop = T) %>% 
      format_spreadsheet(white_spacer, row = even_empty_indices, column = 1:9, loop = T) %>% 
      format_spreadsheet(grand_total_right_corner, row = xlsx_last_instances, column = 9) %>% 
      format_spreadsheet(grand_total_box, row = single_length_indices, column = 9) %>% 
      format_spreadsheet(client_separator, row = sandwiched_indices, column = 1:14, loop = T) %>% 
      format_spreadsheet(red_text, row = negative_indices, column = 13, stack = T) %>% 
      format_spreadsheet(currency_type, row = 2:(nrow(restructured_media) + 1), column = c(5:9, 13), stack = T, loop = T)
    
    ### Adds the Raw Data to the second Tab
    addWorksheet(wb = media_xlsx, sheetName = 'Raw Data', gridLines = T)
    writeData(wb = media_xlsx, sheet = 2, x = raw_media)
    
    ### Outputs a download button
    output$finish_merge <- renderUI(successDownloadButton('download_merge', 'Download'))
    
    ### Outputs the xlsx file into the global
    media_xlsx <<- media_xlsx
    
    ### Sends an alert stating that the task is done
    sendSweetAlert(session = session,
                   title = "Finished merging PDF scrape with raw data!",
                   text = "Excel file is ready for download",
                   type = 'success')
    
  })
  
  ### Excel file download handler
  output$download_merge <- downloadHandler(
    
    ### This name will likely be changed to something more 
    ### custom, but it's not my greatest priority TBH
    filename = 'Media Billing.xlsx', 
    
    ### Allows the file to be handled as a download
    content = function(file) {
      
      ### Saves the workbook to a temp directory
      saveWorkbook(media_xlsx, file = file)
      
    }
    
  )
  
}

### Loads the app
shinyApp(ui = ui, server = server)