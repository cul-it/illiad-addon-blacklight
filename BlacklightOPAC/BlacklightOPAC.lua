local settings = {}
settings.OpacUrl = GetSetting("OPACURL");
local catalog_tou_url = settings.OpacUrl .. "/catalog/tou";
-- need this temporarily.
--local catalog_tou_url = "http://newcatalog5.library.cornell.edu" .. "/catalog/tou";
-----------------
--require "luanet"; -- do not need to do this -- already done by atlas lua add on environment?
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
opacForm.TouInfo = nil;
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
  interfaceMngr = GetInterfaceManager();
  -- Create a form
  opacForm.Form = interfaceMngr:CreateForm("Blacklight OPAC Search", "Script");
  -- Add a browser
  opacForm.Browser=opacForm.Form:CreateBrowser("Blacklight OPAC Search","Blacklight Browser", "Blacklight OPAC Search");
  -- Since we didn't create a ribbon explicitly before creating our browser, it will have created one using the name we passed the CreateBrowser method.  We can retrieve that one and add our buttons to it.
  opacForm.RibbonPage = opacForm.Form:GetRibbonPage("Blacklight OPAC Search");
  -- Create the search and import buttons.
  opacForm.RibbonPage:CreateButton("Search Author",GetClientImage("Search32"),"SearchAuthor","Search");
  opacForm.RibbonPage:CreateButton("Search Keyword",GetClientImage("Search32"),"SearchKeyword","Search");
  opacForm.RibbonPage:CreateButton("Search Title",GetClientImage("Search32"),"SearchTitle", "Search");
  opacForm.RibbonPage:CreateButton("Import Info",GetClientImage("ImportData32"),"ImportInfo", "Import");
  opacForm.RibbonPage:CreateButton("Import as E-Resource",GetClientImage("ImportData32"),"ImportElectronic","Import");
  opacForm.RibbonPage:CreateButton("Open New Browser", GetClientImage("Web32"), "OpenInDefaultBrowser", "Utility");
  opacForm.Browser.WebBrowser.ScriptErrorsSuppressed = true
  opacForm.TouInfo = opacForm.Form:CreateMemoEdit("TOU Info", "TOUInfo");
  opacForm.TouInfo.Value = "Fill in later"; 
  opacForm.JournalInfo = opacForm.Form:CreateMemoEdit("Journal Info", "JournalInfo");
  processType = GetFieldValue("Transaction", "ProcessType");
  opacForm.Form:LoadLayout("BlacklightOPACBorrowlayout.xml");
  opacForm.TouInfo.Value = "Fill in later:" .. processType; 
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
  Log("SearchTerm = " .. searchTerm );
  Log("**** in OPAC LOADED SearchCode = " .. searchCode );
  opacForm.Browser:SetFormValue("search-form", "q", searchTerm);
  opacForm.Browser:SetFormValue("search-form", "search_field", searchCode);
  opacForm.Browser:SubmitForm("search-form");

end

function ImportElectronic()
  local obrowser = opacForm.Browser.WebBrowser;
  local doc_id = string.match(obrowser.DocumentText, "/catalog/(%d+)/citation");
  SetFieldValue("Transaction", "Location", "Olin LIbrary");
  SetFieldValue("Transaction", "CallNumber", "*Networked Resource");
  Log("Blacklight OPAC WebBrowser docid: " .. doc_id);
  SetFieldValue("Transaction","ItemInfo5",settings.OpacUrl .. "/catalog/" .. doc_id);
  Clipboard.SetText(settings.OpacUrl .. "/catalog/" .. doc_id);
  ExecuteCommand("SwitchTab", {"Detail"});
end

function ImportInfo()
  local obrowser = opacForm.Browser.WebBrowser;
  local document = obrowser.Document;
  local url = obrowser.Url;
  local locstr = "";
  local calstr = "";
  local doc_id = string.match(obrowser.DocumentText, "/catalog/(%d+)/citation");
  if doc_id == nil then
    -- Import Info btn was clicked on results page, not on item detail page, so skip import
    return;
  end
  -- the first table should not be the rare table.
  -- local detailsTable = opacForm.Browser:GetElementByCollectionIndex(document:GetElementsByTagName("table"), 0);
  local divs = document:GetElementsByTagName("div");
  -- find the holding div
  if divs == nil then
    return false;
  end
  for i=0,divs.Count - 1 do
    local elem = opacForm.Browser:GetElementByCollectionIndex(divs, i);
    -- if elem.ParentNode ~= nil then
      if elem:GetAttribute("className")=="holding" then
        local holdRows = elem.Children; 
        Log("Blacklight OPAC Found " .. holdRows.Count .. " divs.");
        Log("Blacklight OPAC Found text:'" .. elem.InnerText .. "'.");
        local loctd = opacForm.Browser:GetElementByCollectionIndex(holdRows, 0);
        local caltd = opacForm.Browser:GetElementByCollectionIndex(holdRows, 1);
        Log("Blacklight OPAC loc Text: " .. loctd.InnerText);
        Log("Blacklight OPAC cal Text: " .. caltd.InnerText);
        locstr=string.sub(loctd.InnerText,1,string.len(loctd.InnerText)-8);
        calstr=string.sub(caltd.InnerText,1,string.len(caltd.InnerText));
	      break
      end 
    -- end
  end
  --
  -- in the first holding div, there are two divs, class location, class call-number
  Log("Blacklight OPAC locstr Text: " .. locstr);
  Log("Blacklight OPAC calstr Text: " .. calstr);
  SetFieldValue("Transaction", "Location", locstr);
  SetFieldValue("Transaction", "CallNumber", calstr);
  Log("Blacklight OPAC WebBrowser docid: " .. doc_id);
  Clipboard.SetText(settings.OpacUrl .. "/catalog/" .. doc_id);
  SetFieldValue("Transaction","ItemInfo5",settings.OpacUrl .. "/catalog/" .. doc_id);
  ExecuteCommand("SwitchTab", {"Detail"});
end

function Log(entry)
  if debugEnabled then 
    LogDebug("----- " .. entry .. " -----");
  end
end

