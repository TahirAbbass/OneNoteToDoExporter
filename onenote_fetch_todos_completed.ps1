
# powershell -ExecutionPolicy Bypass -File .\test11.ps1

$onenote = New-Object -ComObject OneNote.Application
Write-Host "OneNote Application connected successfully."

# Your exact target structure
$targetNotebook = "xFlowResearch NoteBook"
$targetSection  = "ToDo"
$targetPage     = "6. To Do List - Daily & Lifelong Meaningful Goals"

# 1. Pull the entire structural map of your OneNote app
$hierarchyXml = ""
$onenote.GetHierarchy("", 4, [ref]$hierarchyXml)

[xml]$hierarchyDoc = $hierarchyXml
$ns = New-Object System.Xml.XmlNamespaceManager($hierarchyDoc.NameTable)
$ns.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")

# 2. Search deep into the map for your explicit Notebook -> Section -> Page path
$xpath = "//one:Notebook[@name='$targetNotebook']//one:Section[@name='$targetSection']//one:Page[@name='$targetPage']"
$pageNode = $hierarchyDoc.SelectSingleNode($xpath, $ns)

if ($pageNode) {
    $pageId = $pageNode.ID
    $sectionId = $pageNode.ParentNode.ID  # Grab parent section ID to host the new page
    Write-Host "Target page found. Source Page ID: $pageId"

    # 3. Extract the raw layout XML for this Page
    $pageContentXml = ""
    $onenote.GetPageContent($pageId, [ref]$pageContentXml, 0)

    [xml]$pageDoc = $pageContentXml

    # 4. FIXED: Enforce strict tag index screening to ignore 'Important' and other non-todo tags
    function Prune-NonCompletedNodes($node, $ns) {
        # Treat Object Elements (OEs) specifically
        if ($node.LocalName -eq "OE") {
            # Only match tags belonging to the "To Do" index definition (index='0') that are marked complete
            $isCompleted = $node.SelectSingleNode("one:Tag[@index='0' and @completed='true']", $ns) -ne $null
            $hasCompletedDescendant = $node.SelectSingleNode(".//one:Tag[@index='0' and @completed='true']", $ns) -ne $null
            
            # If this element is not a completed To-Do AND has no completed To-Do children, delete it entirely
            if (-not ($isCompleted -or $hasCompletedDescendant)) {
                [void]$node.ParentNode.RemoveChild($node)
                return
            }
        }

        # Iterating safely through a snapshot array of children
        if ($node.HasChildNodes) {
            $children = @($node.ChildNodes)
            foreach ($child in $children) {
                Prune-NonCompletedNodes $child $ns
            }
        }

        # Post-cleanup: Remove list wrappers if all their children were pruned
        if ($node.LocalName -eq "OEChildren" -and -not $node.HasChildNodes) {
            [void]$node.ParentNode.RemoveChild($node)
        }
    }

    # 5. Flatten and convert all Tables into standard list layouts
    Write-Host "Unwrapping tables and converting cell contents into standard lists..."
    $tables = @($pageDoc.SelectNodes("//one:Table", $ns))
    foreach ($table in $tables) {
        $parentOE = $table.ParentNode
        $cellOEs = @($table.SelectNodes("./one:Row/one:Cell/one:OEChildren/one:OE", $ns))
        
        $oeChildren = $parentOE.SelectSingleNode("one:OEChildren", $ns)
        if ($null -eq $oeChildren) {
            $oeChildren = $pageDoc.CreateElement("one", "OEChildren", $ns.LookupNamespace("one"))
            [void]$parentOE.AppendChild($oeChildren)
        }
        
        foreach ($oe in $cellOEs) {
            [void]$oeChildren.AppendChild($oe)
        }
        [void]$parentOE.RemoveChild($table)
    }

    # 6. Execute tree pruning on all Outlines
    Write-Host "Filtering completed tasks while maintaining list hierarchy..."
    $outlineNodes = @($pageDoc.SelectNodes("//one:Outline", $ns))
    foreach ($outline in $outlineNodes) {
        Prune-NonCompletedNodes $outline $ns
    }

    # 6b. Consolidate all remaining content into a Single Master Outline positioned at the top
    Write-Host "Consolidating all extracted lists into a single top-level outline block..."
    $oneNsUri = $ns.LookupNamespace("one")
    
    $masterOutline = $pageDoc.CreateElement("one", "Outline", $oneNsUri)
    
    $position = $pageDoc.CreateElement("one", "Position", $oneNsUri)
    $position.SetAttribute("x", "36.0")
    $position.SetAttribute("y", "90.0")
    $position.SetAttribute("z", "0")
    [void]$masterOutline.AppendChild($position)
    
    $size = $pageDoc.CreateElement("one", "Size", $oneNsUri)
    $size.SetAttribute("width", "550.0")
    $size.SetAttribute("height", "0.0")
    [void]$masterOutline.AppendChild($size)
    
    $masterOEChildren = $pageDoc.CreateElement("one", "OEChildren", $oneNsUri)
    
    $survivingOEs = @($pageDoc.SelectNodes("//one:Outline/one:OEChildren/one:OE", $ns))
    foreach ($oe in $survivingOEs) {
        [void]$masterOEChildren.AppendChild($oe)
    }
    [void]$masterOutline.AppendChild($masterOEChildren)
    
    $allOldOutlines = @($pageDoc.SelectNodes("//one:Outline", $ns))
    foreach ($oldOutline in $allOldOutlines) {
        [void]$oldOutline.ParentNode.RemoveChild($oldOutline)
    }
    
    if ($masterOEChildren.HasChildNodes) {
        [void]$pageDoc.DocumentElement.AppendChild($masterOutline)
    }

    # 7. Strip internal objectIDs so OneNote assigns fresh structural tags on insertion
    $elementsWithObjectId = $pageDoc.SelectNodes("//*[@objectID]", $ns)
    foreach ($elem in $elementsWithObjectId) {
        [void]$elem.Attributes.RemoveNamedItem("objectID")
    }

    # 8. Provision a brand new temporary page inside the current section
    $newPageId = ""
    $onenote.CreateNewPage($sectionId, [ref]$newPageId, 0)
    Write-Host "New temporary page provisioned successfully."

    # 9. Re-target the pruned XML tree to point to your new page ID
    $pageDoc.DocumentElement.SetAttribute("ID", $newPageId)
    
    # Update title node text safely
    $titleTextNode = $pageDoc.SelectSingleNode("//one:Title//one:T", $ns)
    if ($titleTextNode) {
        $titleTextNode.InnerText = "Extracted Completed Tasks - Copy"
    }

    # Optional: Save local diagnostic output copy
    $pageDoc.Save("D:\data\code_works\projects\01_onenote_fetch_todos\result.xml")

    # 10. Structured error handling for the synchronization phase
    try {
        $onenote.UpdatePageContent($pageDoc.OuterXml)
        Write-Host "Successfully synchronized filtered tasks into the new page structure." -ForegroundColor Green

        # 11. Bring focus directly to your new page 
        try {
            $onenote.Windows.CurrentWindow.NavigateTo($newPageId)
            Write-Host "Switched window focus to your new page."
        } catch {
            Write-Host "Page created successfully, but focus transition skipped (OneNote window lack of focus)." -ForegroundColor Yellow
        }
    } 
    catch {
        Write-Host "Critical Error updating page contents: $_" -ForegroundColor Red
    }

} else {
    Write-Host "Error: Could not locate Notebook '$targetNotebook' -> Section '$targetSection' -> Page '$targetPage'." -ForegroundColor Red
}