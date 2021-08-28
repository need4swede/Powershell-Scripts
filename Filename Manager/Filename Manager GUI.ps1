#Set-StrictMode -Version 2.0

clear-host

# Main Function
function Main {
    
    # Navigate to desired directory where the files you want modified are stored
    #
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = [System.Environment+SpecialFolder]'MyComputer'
    $browse.ShowNewFolderButton = $true
    $browse.Description = "Locate the folder with the files you want to rename"

    $loop = $true
    while($loop)
    {
        # Once a folder is selected
        #
        if ($browse.ShowDialog() -eq "OK")
        {
            $loop = $false
            $run_define = $true

            # Navigate to selected directory
            #
            cd $browse.SelectedPath

            # MAIN MENU
            #
            do{

            # TASK MENU
            #
            function Define-Task {
                $run_define = $false

                $task_type = read-host "What would you like to do?
                `r`n1. Add Prefix
                `r`n2. Add Suffix
                `r`n3. Remove Text`n
                `r`nSelect Task"
                
                # TASK SELECTION
                #
                $task_loop = $true
                while($task_loop)
                {
                    # ADD PREFIX
                    # Lists filetypes in chosen directory and asks user to add a prefix to the names of those files
                    if ($task_type -eq "1")
                    {
                        $task_loop = $false
                        function add-prefix {
                            $file_ext = read-host "Enter File Type"
                            $file_ext = $file_ext -replace '(^\s+|\s+$)','' -replace '\s+',' '

                            $get_child = get-childitem "*.$file_ext"
                            $add_prefix = read-host "Add the following prefix to all '$file_ext' files"
                            $add_prefix = $add_prefix -replace '(^\s+|\s+$)','' -replace '\s+',' '
                            $get_child | foreach { Rename-Item -Path $_.Name -NewName "$add_prefix-$($_.Name)" }
                        
                        } add-prefix
                        return
                    }

                    # ADD SUFFIX
                    # Lists filetypes in chosen directory and asks user to add a suffix to the names of those files
                    if ($task_type -eq "2")
                    {
                        $task_loop = $false
                        function add-suffix {
                            $file_ext = read-host "Enter File Type"
                            $file_ext = $file_ext -replace '(^\s+|\s+$)','' -replace '\s+',' '

                            $get_child = get-childitem "*.$file_ext"
                            $add_suffix = read-host "Add the following suffix to all '$file_ext' files"
                            $add_suffix = $add_suffix -replace '(^\s+|\s+$)','' -replace '\s+',' '
                            $get_child | foreach { Rename-Item –path $_.Fullname –Newname ( $_.basename + "-" + $add_suffix + $_.extension) }
                        
                        } add-suffix
                        return
                    }

                    # REMOVE TEXT
                    # Lists filetypes in chosen directory and asks user what text they want removed from the names of those files
                    if ($task_type -eq "3")
                    {
                        $task_loop = $false
                        function remove-text {
		                    $file_ext = read-host "Enter File Type"
                            $file_ext = $file_ext -replace '(^\s+|\s+$)','' -replace '\s+',' '

                            $get_child = get-childitem "*.$file_ext"
                            $remove_text = read-host "Remove the following text from all '$file_ext' files"
                            $remove_text = $remove_text -replace '(^\s+|\s+$)','' -replace '\s+',' '
                            $get_child | foreach { rename-item $_ $_.Name.Replace("$remove_text", "") }
                        } remove-text
                        return
                    }
                }
              } Define-Task
              # TASK COMPLETE
              #
              $its_over = $true}until($its_over)
        }
        # 'Cancel' or closing window exits program
        #
        else{return}
    }
    
    $browse.SelectedPath
    $browse.Dispose()
} Main