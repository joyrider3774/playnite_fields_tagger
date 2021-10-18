function GetMainMenuItems
{
    param(
        $getMainMenuItemsArgs
    )

    $menuItem1 = New-Object Playnite.SDK.Plugins.ScriptMainMenuItem
    $menuItem1.Description = ([Playnite.SDK.ResourceProvider]::GetString("LOCGame_Field_Tagger_MenuItemOpenMenuDescription"))
    $menuItem1.FunctionName = "OpenMenu"
    $menuItem1.MenuSection = "@Fields Tagger"

    return $menuItem1
}


function OpenMenu
{
    param(
        $scriptMainMenuItemActionArgs
    )

    # Load assemblies
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName PresentationFramework
    
    # Set Xaml
    [xml]$Xaml = @"
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <Grid.Resources>
        <Style TargetType="TextBlock" BasedOn="{StaticResource BaseTextBlockStyle}" />
    </Grid.Resources>

    <StackPanel Margin="20">
        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
            <TextBlock TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_GameSelection}" VerticalAlignment="Center" MinWidth="140"/>
            <ComboBox Name="CbGameSelection" SelectedIndex="0" MinHeight="25" MinWidth="200" VerticalAlignment="Center" Margin="10,0,0,0">
                <ComboBoxItem Content="{DynamicResource LOCGame_Field_Tagger_Allgames}" HorizontalAlignment="Stretch"/>
                <ComboBoxItem Content="{DynamicResource LOCGame_Field_Tagger_Selectedgames}" HorizontalAlignment="Stretch"/>
            </ComboBox>
        </StackPanel>

        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
            <TextBlock TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_Fieldselection}" VerticalAlignment="Center" MinWidth="140"/>
            <ComboBox Name="CbField" SelectedIndex="0" MinHeight="25" MinWidth="200" VerticalAlignment="Center" Margin="10,0,0,0">
                <ComboBoxItem Content="Description" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="GameStartedScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="IncludeLibraryPluginAction" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="InstallDirectory" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="IsCustomGame" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="Manual" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="Notes" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="PostScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="PreScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="SortingName" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="UseGlobalGameStartedScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="UseGlobalPostScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="UseGlobalPreScript" HorizontalAlignment="Stretch"/>
				<ComboBoxItem Content="Version" HorizontalAlignment="Stretch"/>
            </ComboBox>
        </StackPanel>
        <TabControl Name="ControlTools" HorizontalAlignment="Left" MinHeight="220" Margin="0,0,0,15">
            <TabItem Header="{DynamicResource LOCGame_Field_Tagger_MissingFields}">
                <StackPanel>
                    <TextBlock Margin="0,15,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_Description}" FontWeight="Bold"/>
                    <TextBlock Margin="0,0,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_ToolMissing}"/>
                </StackPanel>
            </TabItem>
			<TabItem Header="{DynamicResource LOCGame_Field_Tagger_HasFields}">
                <StackPanel>
                    <TextBlock Margin="0,15,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_Description}" FontWeight="Bold"/>
                    <TextBlock Margin="0,0,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_ToolHave}"/>
                </StackPanel>
            </TabItem>
			<TabItem Header="{DynamicResource LOCGame_Field_Tagger_RemoveTags}">
                <StackPanel>
                    <TextBlock Margin="0,15,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_Description}" FontWeight="Bold"/>
                    <TextBlock Margin="0,0,0,15" TextWrapping="Wrap" Text="{DynamicResource LOCGame_Field_Tagger_ToolRemove}"/>
                </StackPanel>
            </TabItem>
        </TabControl>
        <Button Content="{DynamicResource LOCGame_Field_Tagger_UpdateTags}" HorizontalAlignment="Center" Margin="0,0,0,15" Name="ButtonUpdateTags" IsDefault="True"/>
    </StackPanel>
</Grid>
"@

    # Load the xaml for controls
    $XMLReader = [System.Xml.XmlNodeReader]::New($Xaml)
    $XMLForm = [Windows.Markup.XamlReader]::Load($XMLReader)

    # Make variables for each control
    $Xaml.FirstChild.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name $_.Name -Value $XMLForm.FindName($_.Name) }

    # Set Window creation options
    $windowCreationOptions = New-Object Playnite.SDK.WindowCreationOptions
    $windowCreationOptions.ShowCloseButton = $true
    $windowCreationOptions.ShowMaximizeButton = $False
    $windowCreationOptions.ShowMinimizeButton = $False

    # Create window
    $window = $PlayniteApi.Dialogs.CreateWindow($windowCreationOptions)
    $window.Content = $XMLForm
    $window.Width = 620
    $window.Height = 460
    $window.Title = "Fields Tagger"
    $window.WindowStartupLocation = "CenterScreen"

    # Handler for pressing "Add Tags" button
    $ButtonUpdateTags.Add_Click(
    {
        # Get the variables from the controls
        $gameSelection = $CbGameSelection.SelectedIndex
        $toolSelection = $ControlTools.SelectedIndex

        # Set GameDatabase
        switch ($gameSelection) {
            0 {
                $gameDatabase = $PlayniteApi.Database.Games
                $__logger.Info("Fields Tagger - Game Selection: `"AllGames`"")
            }
            1 {
                $gameDatabase = $PlayniteApi.MainView.SelectedGames
                $__logger.Info("Fields Tagger - Game Selection: `"SelectedGames`"")
            }
        }

        # Set Field Type
       $Field = $CbField.SelectedValue.Content
       $__logger.Info("Fields Tagger - Field Selection: `"$Field`"")

        # Set Tool
        switch ($toolSelection) {
            0 { # Tool #0: Missing Field

                $__logger.Info("Fields Tagger - Tool Selection: `"Missing Field`"")
                
                # Start Fields Tagger function
                $__logger.Info("Fields Tagger - Starting Function with parameters `"$Field`"")
                Invoke-MissingFieldTagger $gameDatabase $Field
            }
			1 { # Tool #1: Have Field

                $__logger.Info("Fields Tagger - Tool Selection: `"Have Field`"")
                
                # Start Fields Tagger function
                $__logger.Info("Fields Tagger - Starting Function with parameters `"$Field`"")
                Invoke-HasFieldTagger $gameDatabase $Field
            }
			2 { # Tool #2: Remove Tags

                $__logger.Info("Fields Tagger - Tool Selection: `"Have Field`"")
                
                # Start Fields Tagger function
                $__logger.Info("Fields Tagger - Starting Function with parameters `"$Field`"")
                Invoke-RemoveTags $gameDatabase $Field
            }
        }
    })

    # Show Window
    $__logger.Info("Fields Tagger - Opening Window.")
    $window.ShowDialog()
    $__logger.Info("Fields Tagger - Window closed.")
}

function Invoke-MissingFieldTagger
{
    param (
        $gameDatabase, 
        $Field
    )
    
    # Create "No Field" tag
    $tagNoFieldName = "No Field: " + $Field
    $tagNoField = $PlayniteApi.Database.tags.Add($tagNoFieldName)
    $tagNoFieldIds = $tagNoField.Id

    
    foreach ($game in $gameDatabase) {

        $FieldValue = $game.$Field
        if (($null -eq $FieldValue) -or ($FieldValue -eq $false) -or (($FieldValue) -and ($FieldValue -eq "")))
        {
            Add-TagToGame $game $tagNoFieldIds
        }
        else
        {
            Remove-TagFromGame $game $tagNoFieldIds
        }
    }
    
    # Generate results of missing Field in selection
    $GamesNoFieldSelection = $gameDatabase | Where-Object {$_.TagIds -contains $tagNoFieldIds.Guid}
    $results = ([Playnite.SDK.ResourceProvider]::GetString("LOCGame_Field_Tagger_Results1Message") -f $($gameDatabase.count.ToString()), $Field, $($GamesNoFieldSelection.Count.ToString()) )

    $__logger.Info("Fields Tagger - $($results -replace "`n", ', ')")
    $PlayniteApi.Dialogs.ShowMessage($results, "Fields Tagger")
}

function Invoke-HasFieldTagger
{
    param (
        $gameDatabase, 
        $Field
    )
    
    # Create "No Field" tag
    $tagFieldName = "Has Field: " + $Field
    $tagField = $PlayniteApi.Database.tags.Add($tagFieldName)
    $tagFieldIds = $tagField.Id

    
    foreach ($game in $gameDatabase) {

        $FieldValue = $game.$Field
        if (($FieldValue) -and ($FieldValue -ne ""))
        {
            Add-TagToGame $game $tagFieldIds
        }
        else
        {
            Remove-TagFromGame $game $tagFieldIds
        }
    }
    
    # Generate results of missing Field in selection
    $GamesNoFieldSelection = $gameDatabase | Where-Object {$_.TagIds -contains $tagFieldIds.Guid}
    $results = ([Playnite.SDK.ResourceProvider]::GetString("LOCGame_Field_Tagger_Results2Message") -f $($gameDatabase.count.ToString()), $Field, $($GamesNoFieldSelection.Count.ToString()) )

    $__logger.Info("Fields Tagger - $($results -replace "`n", ', ')")
    $PlayniteApi.Dialogs.ShowMessage($results, "Fields Tagger")
}

function Invoke-RemoveTags
{
    param (
        $gameDatabase, 
        $Field
    )
    
    $tagFieldName = "Has Field: " + $Field
    $tagField = $PlayniteApi.Database.tags.Add($tagFieldName)
    $tagFieldIds = $tagField.Id

    $tagNoFieldName = "No Field: " + $Field
    $tagNoField = $PlayniteApi.Database.tags.Add($tagNoFieldName)
    $tagNoFieldIds = $tagNoField.Id
	
    foreach ($game in $gameDatabase) {
         Remove-TagFromGame $game $tagNoFieldIds
		 Remove-TagFromGame $game $tagFieldIds
    }
	
    $tagNoFieldNameSelection = $gameDatabase | Where-Object {$_.TagIds -contains $tagNoFieldIds.Guid}
	$tagFieldNameSelection = $gameDatabase | Where-Object {$_.TagIds -contains $tagFieldIds.Guid}
    $tagNoFieldNameAll = $PlayniteApi.Database.Games | Where-Object {$_.TagIds -contains $tagNoFieldIds.Guid}
	$tagFieldNameAll = $PlayniteApi.Database.Games | Where-Object {$_.TagIds -contains $tagFieldIds.Guid}
	 
    $__logger.Info("Fields Tagger - Games with tag `"$tagNoFieldName`" at finish: Selection $($tagNoFieldNameSelection.count), All $($tagNoFieldNameAll.count)")
	$__logger.Info("Fields Tagger - Games with tag `"$tagFieldName`" at finish: Selection $($tagFieldNameSelection.count), All $($tagFieldNameAll.count)")
          
	# Remove tool tag from database if 0 games have it
	if (($tagNoFieldNameSelection.count -eq 0) -and ($tagNoFieldNameAll.count -eq 0))
	{
		$PlayniteApi.Database.Tags.Remove($tagNoFieldIds)
		$__logger.Info("Fields Tagger - Removed tag `"$tagNoFieldName`" from database")
	}
	
	if (($tagFieldNameSelection.count -eq 0) -and ($tagFieldNameAll.count -eq 0))
	{
		$PlayniteApi.Database.Tags.Remove($tagFieldIds)
		$__logger.Info("Fields Tagger - Removed tag `"$tagFieldName`" from database")
	}

	$results = ([Playnite.SDK.ResourceProvider]::GetString("LOCGame_Field_Tagger_Results3Message") -f $tagNoFieldName, $($tagNoFieldNameSelection.count.ToString()), $($tagNoFieldNameAll.count.ToString()) )
	$results += "`n" + ([Playnite.SDK.ResourceProvider]::GetString("LOCGame_Field_Tagger_Results3Message") -f $tagFieldName, $($tagFieldNameSelection.count.ToString()), $($tagFieldNameAll.count.ToString()) )
   
    $PlayniteApi.Dialogs.ShowMessage($results, "Fields Tagger")
}

function Add-TagToGame
{
    param (
        $game,
        $tagIds
    )

    # Check if game already doesn't have tag
    if ($game.tagIds -notcontains $tagIds)
    {
        # Add tag Id to game
        if ($game.tagIds)
        {
            $game.tagIds += $tagIds
        }
        else
        {
            # Fix in case game has null tagIds
            $game.tagIds = $tagIds
        }
        
        # Update game in database and increase no Field count
        $PlayniteApi.Database.Games.Update($game)
    }
}

function Remove-TagFromGame
{
    param (
        $game,
        $tagIds
    )

    # Check if game already has tag and remove it
    if ($game.tagIds -contains $tagIds)
    {
        $game.tagIds.Remove($tagIds)
        $PlayniteApi.Database.Games.Update($game)
    }
}