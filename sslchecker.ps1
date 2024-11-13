# PowerShell script to display SSL certificate check results and send email

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to display the results in a popup with email input
function Show-ResultsPopup {
    param (
        [string]$Results
    )

    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SSL Certificate Checker Results"
    $form.Size = New-Object System.Drawing.Size(800, 500)
    $form.StartPosition = "CenterScreen"

    # Text box to display results
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Size = New-Object System.Drawing.Size(576, 300)
    $textBox.Location = New-Object System.Drawing.Point(1, 1)
    $textBox.Font = New-Object System.Drawing.Font("Courier New", 10)  # Fixed-width font for better alignment
    $textBox.Text = $Results
    $form.Controls.Add($textBox)

    # Label for email
    $emailLabel = New-Object System.Windows.Forms.Label
    $emailLabel.Text = "Enter Email Address:"
    $emailLabel.Location = New-Object System.Drawing.Point(10, 320)
    $form.Controls.Add($emailLabel)

    # Text box for email input
    $emailBox = New-Object System.Windows.Forms.TextBox
    $emailBox.Size = New-Object System.Drawing.Size(500, 20)
    $emailBox.Location = New-Object System.Drawing.Point(10, 340)
    $form.Controls.Add($emailBox)

    # Button to send email
    $sendButton = New-Object System.Windows.Forms.Button
    $sendButton.Text = "Send Email"
    $sendButton.Location = New-Object System.Drawing.Point(520, 340)
    $sendButton.Add_Click({
        $email = $emailBox.Text
        if ($email -match "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$") {
            # Call the Python script with the email argument and pass the email address from the input field
            & python.exe "C:\Users\xxx\xxx\Script\sslcheck.py" -f "C:\Users\xxx\xxx\Script\domains.txt" -e $email  #use absolute path of the the all file
            [System.Windows.Forms.MessageBox]::Show("Email sent to $email.")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid email format.")
        }
    })
    $form.Controls.Add($sendButton)

    # Show the form
    $form.ShowDialog()
}

# Step 1: Run the Python script to get SSL check results
# Capture the output from the Python script and assign it to $sslResults
$sslResults = & python.exe "C:\Users\xxx\xxx\Script\sslcheck.py" -f "C:\Users\xxx\xxx\Script\domains.txt"

# Step 2: Display the results in a popup form
Show-ResultsPopup -Results $sslResults
