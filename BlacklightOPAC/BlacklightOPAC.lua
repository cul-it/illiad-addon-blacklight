local settings = {}
settings.OpacUrl = GetSetting("OPACURL");

luanet.load_assembly("System");
luanet.load_assembly("System.Windows");
luanet.load_assembly("System.Windows.Forms");

local Clipboard  = luanet.import_type("System.Windows.Forms.Clipboard");
local WebClient  = luanet.import_type("System.Net.WebClient")
local wclient = WebClient();

local interfaceMngr = nil;
local opacForm = {};
opacForm.Form = nil;
opacForm.RibbonPage = nil;
opacForm.Browser = nil;
opacForm.JournalInfo = nil;

local searchTerm = nil;
local searchCode = nil;
local processType = nil;

local debugEnabled = true;

function stripc(str,chrs)
  local s = str:gsub("["..chrs:gsub("%W","%%%1").."]",'');
  return s
end

function Init()
  LogDebug("\n");
  LogDebug("\n\n==================== BLACKLIGHT OPAC INIT ====================\n");
  interfaceMngr = GetInterfaceManager();
  -- Create a form
  opacForm.Form = interfaceMngr:CreateForm("Blacklight OPAC Search", "Script");
  opacForm.JournalInfo = opacForm.Form:CreateMemoEdit("Journal Info", "JournalInfo");

  -- Add a browser, WebView2 is Microsoft Edge based browser control supports bootstrap 5
  opacForm.Browser = opacForm.Form:CreateBrowser("Blacklight OPAC Search", "Blacklight Browser", "Blacklight OPAC Search", "WebView2");
  
  -- Turn off accelerator keys so the browser doesn't hijack the Tab focus away from the addon tabs, a known issue with webview2 browser
  local nativeBrowser = opacForm.Browser.Control or opacForm.Browser.WebBrowser;
  if nativeBrowser and nativeBrowser.CoreWebView2 then
    nativeBrowser.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
  end

  -- Since we didn't create a ribbon explicitly before creating our browser, it will have created one using the name we passed the CreateBrowser method.  We can retrieve that one and add our buttons to it.
  opacForm.RibbonPage = opacForm.Form:GetRibbonPage("Blacklight OPAC Search");

  -- Create the search and import buttons.
  opacForm.RibbonPage:CreateButton("Search Author",GetClientImage("Search32"),"SearchAuthor","Search");
  opacForm.RibbonPage:CreateButton("Search Keyword",GetClientImage("Search32"),"SearchKeyword","Search");
  opacForm.RibbonPage:CreateButton("Search Title",GetClientImage("Search32"),"SearchTitle", "Search");
  opacForm.RibbonPage:CreateButton("Import Info",GetClientImage("ImportData32"),"ImportInfo", "Import");
  opacForm.RibbonPage:CreateButton("Import as E-Resource",GetClientImage("ImportData32"),"ImportElectronic","Import");
  opacForm.RibbonPage:CreateButton("Open New Browser", GetClientImage("Web32"), "OpenInDefaultBrowser", "Utility");

  processType = GetFieldValue("Transaction", "ProcessType");
  opacForm.Form:LoadLayout("webview2formlayout.xml");

  opacForm.Form:Show();
  SearchTitle();
end

-- Open New Browser button 
function OpenInDefaultBrowser()
  local currentUrl = opacForm.Browser.Address;
  if (currentUrl and currentUrl ~= "") then
      LogDebug("Opening Browser URL in default browser: " .. currentUrl);
      os.execute('start "" "' .. currentUrl .. '"');
  end
end

function SearchKeyword()
  if GetFieldValue("Transaction", "RequestType") == "Loan" then
    searchTerm = GetFieldValue("Transaction", "LoanTitle");
    searchCode = "all_fields";
  else
    searchTerm = GetFieldValue("Transaction", "PhotoJournalTitle");
    journalTitle = GetFieldValue("Transaction", "PhotoArticleTitle");
    journalAuthor = GetFieldValue("Transaction", "PhotoArticleAuthor");
    journalYear = GetFieldValue("Transaction", "PhotoJournalYear");
    journalVolume = GetFieldValue("Transaction", "PhotoJournalVolume");
    journalPages = GetFieldValue("Transaction", "PhotoJournalInclusivePages");
    journalIssue = GetFieldValue("Transaction", "PhotoJournalIssue");
    if journalIssue == nil then
      journalIssue = "";
    end
    searchTerm = stripc(searchTerm,"/:");
    opacForm.JournalInfo.Value = searchTerm .. " " .. processType .. " Article: " 
      .. journalTitle .. " Year: " .. journalYear 
      .. journalAuthor .. " Volume: " .. journalVolume 
      ..  " Issue: " .. journalIssue 
      ..  " Pages: " .. journalPages; 
    searchCode = "all_fields";
  end
  searchTerm = stripc(searchTerm,"/:");
  opacForm.Browser:RegisterPageHandler("formExists", "search-form", "OPACLoaded", false);
  opacForm.Browser:Navigate(settings.OpacUrl);	
end

function SearchTitle()
  if GetFieldValue("Transaction", "RequestType") == "Loan" then
    searchTerm = GetFieldValue("Transaction", "LoanTitle");
    opacForm.JournalInfo.Value = searchTerm; 
    searchCode = "title";
  else
    searchTerm = GetFieldValue("Transaction", "PhotoJournalTitle");
    journalTitle = GetFieldValue("Transaction", "PhotoArticleTitle");
    journalYear = GetFieldValue("Transaction", "PhotoJournalYear");
    journalAuthor = GetFieldValue("Transaction", "PhotoArticleAuthor");
    journalVolume = GetFieldValue("Transaction", "PhotoJournalVolume");
    journalPages = GetFieldValue("Transaction", "PhotoJournalInclusivePages");
    journalIssue = GetFieldValue("Transaction", "PhotoJournalIssue");
    if journalIssue == nil then
      journalIssue = "";
    end
    opacForm.JournalInfo.Value = searchTerm .. " " .. processType .. " Article: " 
      .. journalTitle .. " Year: " .. journalYear 
      .. journalAuthor .. " Volume: " .. journalVolume 
      ..  " Issue: " .. journalIssue 
      ..  " Pages: " .. journalPages; 
    searchCode = "title";
    --searchCode = "journal title";
  end
  searchTerm = stripc(searchTerm,"/:");
  opacForm.Browser:RegisterPageHandler("formExists", "search-form", "OPACLoaded", false);
  opacForm.Browser:Navigate(settings.OpacUrl);
end

-- Called when the 'Search Author' button is clicked
-- fetch author information from the transaction and then search the catalog for more items by that author
function SearchAuthor()
  if GetFieldValue("Transaction", "RequestType") == "Loan" then
    searchTerm = GetFieldValue("Transaction", "LoanAuthor");
  else
    searchTerm = GetFieldValue("Transaction", "PhotoArticleAuthor");
    opacForm.JournalInfo.Value = searchTerm .. " " .. processType ; 
  end
  searchCode = "author";
  searchTerm = stripc(searchTerm,"/:");
  opacForm.Browser:RegisterPageHandler("formExists", "search-form", "OPACLoaded", false);
  opacForm.Browser:Navigate(settings.OpacUrl);  
end

function OPACLoaded()
  Log("**** in OPAC LOADED SearchCode = " .. searchCode );
  local jsTerm = searchTerm:gsub("\\", "\\\\"):gsub("'", "\\'");
  local script = [[
    (function() {
      var form = document.getElementById('search-form') || document.forms['search-form'];
      if (!form) return 'no form';
      form.elements['q'].value = ']] .. jsTerm .. [[';
      form.elements['search_field'].value = ']] .. searchCode .. [[';
      form.submit();
      return 'submitted';
    })();
  ]];
  opacForm.Browser:ExecuteScript(script);
end

function ImportElectronic()
  local doc_id = string.match(tostring(opacForm.Browser.Address), "/catalog/(%d+)");
  if doc_id == nil then return; end
  SetFieldValue("Transaction", "Location", "Olin Library");
  SetFieldValue("Transaction", "CallNumber", "*Networked Resource");
  Log("Blacklight OPAC WebBrowser docid: " .. doc_id);
  SetFieldValue("Transaction","ItemInfo5",settings.OpacUrl .. "/catalog/" .. doc_id);
  Clipboard.SetText(settings.OpacUrl .. "/catalog/" .. doc_id);
  ExecuteCommand("SwitchTab", {"Detail"});
end


--- Import bibliographic details from the currently loaded Blacklight item page.
---
--- Behavior:
--- 1. Reads the current catalog record id from the browser's URL
--- 2. Fetches the record page HTML over HTTP via WebClient
--- 3. Finds the first element with class "holding".
--- 4. Pulls location and call number from the element's data attributes:
---    - data-location
---    - data-call-number
--- 5. Writes values to the active ILLiad transaction and stores the item URL in
---    ItemInfo5 and the system clipboard.
---
--- Notes:
--- - If the page is not a record detail page (no catalog id in the URL), the function exits.
--- - If the HTTP call fails, the function exits.
--- - If no holding element is found, the function returns false.
--- - If multiple holdings exist, the first one is used.
function ImportInfo()
  local currentUrl = tostring(opacForm.Browser.Address);
  local doc_id = string.match(currentUrl, "/catalog/(%d+)");
  if doc_id == nil then
    LogDebug("ImportInfo: no doc_id in URL - not a record detail page, skipping");
    return;
  end

  -- Fetch the record page HTML via WebClient (ExecuteScript does not return values with webview2 browser)
  local ok, pageHtml = pcall(function()
    return wclient:DownloadString(settings.OpacUrl .. "/catalog/" .. doc_id);
  end);
  if not ok then
    LogDebug("ImportInfo: HTTP fetch failed: " .. tostring(pageHtml));
    return;
  end
  pageHtml = tostring(pageHtml);

  -- Parse first holding's data attributes (note: attribute is 'data-callnumber')
  local locstr, calstr = string.match(pageHtml,
    'class="holding"%s+data%-location="(.-)"%s+data%-callnumber="(.-)"');
  if locstr == nil then
    LogDebug("ImportInfo: no holding found in page HTML");
    return false;
  end

  SetFieldValue("Transaction", "Location", locstr);
  SetFieldValue("Transaction", "CallNumber", calstr);
  Clipboard.SetText(settings.OpacUrl .. "/catalog/" .. doc_id);
  SetFieldValue("Transaction", "ItemInfo5", settings.OpacUrl .. "/catalog/" .. doc_id);
  LogDebug("ImportInfo: fields set, switching to Detail tab");
  ExecuteCommand("SwitchTab", {"Detail"});
end

function Log(entry)
  if debugEnabled then 
    LogDebug("----- " .. entry .. " -----");
  end
end
