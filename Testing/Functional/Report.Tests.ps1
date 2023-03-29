Import-Module Selenium

Describe -Tag "UI","Chrome" -Name "Test Report with <Browser>" -ForEach @(
    @{ Browser = "Chrome"; Driver = Start-SeChrome 2>$null }
    @{ Browser = "Edge"; Driver = Start-SeEdge 2>$null }
){
	BeforeAll {
        $script:url = "file:///C:/Users/crutchfield/source/repos/ScubaGear/Testing/M365BaselineConformance_2023_03_24_08_41_48/BaselineReports.html"
        Enter-SeUrl $script:url -Driver $Driver 2>$null
	}

	It "Toggle Dark Mode" {
        $ToggleCheckbox = Find-SeElement -Driver $Driver -Wait -By XPath "//input[@id='toggle']"
        $ToggleText = Find-SeElement -Driver $Driver -Wait -Id "toggle-text"

        $ToggleCheckbox.Selected | Should -Be $false
        $ToggleText.Text | Should -Be 'Light Mode'

        $ToggleSwitch = Find-SeElement -Driver $Driver -Wait -ClassName "switch"
        Invoke-SeClick -Element $ToggleSwitch

        $ToggleText.Text | Should -Be 'Dark Mode'
        $ToggleCheckbox.Selected | Should -Be $true
	}

    It "Verify Tenant"{
        $TenantDataElement = Find-SeElement -Driver $Driver -Wait -ClassName "tenantdata"
        $TenantDataRows = Find-SeElement -Target $TenantDataElement -By TagName "tr"
        $TenantDataColumns = Find-SeElement -Target $TenantDataRows[1] -By TagName "td"
        $Tenant = $TenantDataColumns[0].Text
        $Tenant | Should -Be "Cybersecurity and Infrastructure Security Agency" -Because $Tenant
    }

    It "Verify  Domain"{
        $TenantDataElement = Find-SeElement -Driver $Driver -Wait -ClassName "tenantdata"
        $TenantDataRows = Find-SeElement -Target $TenantDataElement -By TagName "tr"
        $TenantDataColumns = Find-SeElement -Target $TenantDataRows[1] -By TagName "td"
        $Domain = $TenantDataColumns[1].Text
        $Domain | Should -Be "cisaent.onmicrosoft.com" -Because "Domain is $Domain"
    }

    It "Goto <Product> (<LinkText>) details" -ForEach @(
        @{Product = "aad"; LinkText = "Azure Active Directory"}
        @{Product = "onedrive"; LinkText = "OneDrive for Business"}
    ){
        $DetailLink = Find-SeElement -Driver $Driver -Wait -By LinkText $LinkText
        $DetailLink | Should -Not -BeNullOrEmpty
        Invoke-SeClick -Element $DetailLink

        $ToggleCheckbox = Find-SeElement -Driver $Driver -Wait -By XPath "//input[@id='toggle']"
        $ToggleText = Find-SeElement -Driver $Driver -Wait -Id "toggle-text"

        $ToggleText.Text | Should -Be 'Dark Mode'
        $ToggleCheckbox.Selected | Should -Be $true

        $ToggleSwitch = Find-SeElement -Driver $Driver -Wait -ClassName "switch"
        Invoke-SeClick -Element $ToggleSwitch

        $ToggleText.Text | Should -Be 'Light Mode'
        $ToggleCheckbox.Selected | Should -Be $false
    }

    It "Go Back to main page - Is Dark mode in correct state"{
        Open-SeUrl -Back -Driver $Driver
        $ToggleCheckbox = Find-SeElement -Driver $Driver -Wait -By XPath "//input[@id='toggle']"
        $ToggleText = Find-SeElement -Driver $Driver -Wait -Id "toggle-text"
        $ToggleText.Text | Should -Be 'Light Mode'
        $ToggleCheckbox.Selected | Should -Be $false
    }

	AfterAll {
		Stop-SeDriver -Driver $Driver 2>$null
	}
}
