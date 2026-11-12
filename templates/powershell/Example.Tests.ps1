# Example.Tests.ps1
#
# TEMPLATE FILE: Copy this file to tests/PowerShell/ and customize for your project.
#
# This file demonstrates Pester 5.x test patterns including:
#   - BeforeAll/BeforeEach for test setup
#   - Describe/Context/It block structure
#   - Arrange-Act-Assert (AAA) pattern
#   - Testing integer return codes (v1.0 style)
#   - Testing reference parameters ([ref])
#   - Testing boolean returns (Test-* functions)
#   - Basic mocking with Mock command
#
# Prerequisites:
#   - Pester 5.x: Install-Module -Name Pester -MinimumVersion 5.0 -Force
#
# Usage (from repository root):
#   Invoke-Pester -Path templates/powershell/Example.Tests.ps1 -Output Detailed
#
# For complete Pester documentation: https://pester.dev/docs/quick-start

# BeforeAll runs once before all tests in this file.
# Use this to dot-source the script/module containing the functions under test.
BeforeAll {
    # Example: Dot-source the function file being tested
    # . $PSScriptRoot/../../src/MyFunction.ps1

    # For this template, we define sample functions inline for demonstration.
    # In your actual tests, remove these and dot-source your real functions.

    function Get-ExampleGreeting {
        # .SYNOPSIS
        # Returns a greeting message.
        #
        # .DESCRIPTION
        # Constructs a greeting message for a given name and returns it via a
        # reference parameter. This function demonstrates the v1.0 return code
        # pattern with reference parameters for output.
        #
        # .PARAMETER ReferenceToResult
        # A reference to a string variable that will receive the greeting message.
        #
        # .PARAMETER Name
        # The name to include in the greeting message. Must not be null or
        # whitespace.
        #
        # .EXAMPLE
        # $strGreeting = $null
        # $intResult = Get-ExampleGreeting -ReferenceToResult ([ref]$strGreeting) -Name "World"
        # if ($intResult -eq 0) {
        #     Write-Host $strGreeting  # Outputs: Hello, World!
        # }
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # [int] Status code: 0=success, -1=failure
        #
        # .NOTES
        # Version: 1.0.20260119.0
        # This is a demonstration function for the Pester testing template.
        #
        param(
            [ref]$ReferenceToResult,
            [string]$Name
        )

        if ([string]::IsNullOrWhiteSpace($Name)) {
            return -1
        }

        $ReferenceToResult.Value = "Hello, $Name!"
        return 0
    }

    function Test-IsValidEmail {
        # .SYNOPSIS
        # Tests if a string is a valid email format.
        #
        # .DESCRIPTION
        # Validates whether the provided string matches a basic email format
        # pattern. This function demonstrates the boolean return pattern for
        # Test-* functions where simple true/false validation is needed.
        #
        # .PARAMETER Email
        # The email address string to validate. Must contain an @ symbol with
        # text before and after, and a domain extension.
        #
        # .EXAMPLE
        # $boolIsValid = Test-IsValidEmail -Email "user@example.com"
        # if ($boolIsValid) {
        #     Write-Host "Email is valid"
        # }
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # [bool] $true if valid, $false otherwise
        #
        # .NOTES
        # Version: 1.0.20260119.0
        # This is a demonstration function for the Pester testing template.
        #
        param(
            [string]$Email
        )

        if ([string]::IsNullOrWhiteSpace($Email)) {
            return $false
        }

        return $Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    }

    function Get-ProcessedData {
        # .SYNOPSIS
        # Example function that calls external command (for mocking demo).
        #
        # .DESCRIPTION
        # Reads content from a specified file path and returns it via a reference
        # parameter. This function demonstrates how to test functions that call
        # external commands by using Pester's Mock capability.
        #
        # .PARAMETER ReferenceToResult
        # A reference to a variable that will receive the file content.
        #
        # .PARAMETER Source
        # The file path to read content from.
        #
        # .EXAMPLE
        # $arrContent = $null
        # $intResult = Get-ProcessedData -ReferenceToResult ([ref]$arrContent) -Source "C:\data\input.txt"
        # if ($intResult -eq 0) {
        #     Write-Host "Read $($arrContent.Count) lines"
        # }
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # [int] Status code: 0=success, -1=failure
        #
        # .NOTES
        # Version: 1.0.20260119.0
        # This is a demonstration function for the Pester testing template.
        #
        param(
            [ref]$ReferenceToResult,
            [string]$Source
        )

        # This calls Get-Content which we can mock in tests
        $content = Get-Content -Path $Source -ErrorAction SilentlyContinue

        if ($null -eq $content) {
            return -1
        }

        $ReferenceToResult.Value = $content
        return 0
    }
}

# ==============================================================================
# Testing Functions with Integer Return Codes
# ==============================================================================
# These tests demonstrate the pattern for functions that return:
#   0 = success
#   -1 = failure
#   1-5 = partial success (with additional data)

Describe "Get-ExampleGreeting" {
    # Context blocks group related tests for a specific scenario
    Context "When given a valid name" {
        It "Returns success code 0" {
            # Arrange - Set up test data and expected results
            $strResult = $null
            $strName = "World"

            # Act - Execute the function under test
            $intReturnCode = Get-ExampleGreeting -ReferenceToResult ([ref]$strResult) -Name $strName

            # Assert - Verify the expected outcome
            $intReturnCode | Should -Be 0
        }

        It "Populates the reference parameter with greeting" {
            # Arrange
            $strResult = $null
            $strName = "Pester"

            # Act
            [void](Get-ExampleGreeting -ReferenceToResult ([ref]$strResult) -Name $strName)

            # Assert
            $strResult | Should -Be "Hello, Pester!"
        }
    }

    Context "When given an empty name" {
        It "Returns failure code -1" {
            # Arrange
            $strResult = $null
            $strName = ""

            # Act
            $intReturnCode = Get-ExampleGreeting -ReferenceToResult ([ref]$strResult) -Name $strName

            # Assert
            $intReturnCode | Should -Be -1
        }

        It "Does not populate the reference parameter" {
            # Arrange
            $strResult = $null
            $strName = "   "  # Whitespace only

            # Act
            [void](Get-ExampleGreeting -ReferenceToResult ([ref]$strResult) -Name $strName)

            # Assert
            $strResult | Should -BeNullOrEmpty
        }
    }
}

# ==============================================================================
# Testing Test-* Functions (Boolean Returns)
# ==============================================================================
# Test-* functions that return $true/$false should have tests for both cases.

Describe "Test-IsValidEmail" {
    Context "When email is valid" {
        It "Returns true for standard email format" {
            # Arrange
            $strEmail = "user@example.com"

            # Act
            $boolResult = Test-IsValidEmail -Email $strEmail

            # Assert
            $boolResult | Should -BeTrue
        }

        It "Returns true for email with subdomain" {
            # Arrange
            $strEmail = "user@mail.example.com"

            # Act
            $boolResult = Test-IsValidEmail -Email $strEmail

            # Assert
            $boolResult | Should -BeTrue
        }
    }

    Context "When email is invalid" {
        It "Returns false for empty string" {
            # Arrange
            $strEmail = ""

            # Act
            $boolResult = Test-IsValidEmail -Email $strEmail

            # Assert
            $boolResult | Should -BeFalse
        }

        It "Returns false for missing @ symbol" {
            # Arrange
            $strEmail = "userexample.com"

            # Act
            $boolResult = Test-IsValidEmail -Email $strEmail

            # Assert
            $boolResult | Should -BeFalse
        }

        It "Returns false for missing domain" {
            # Arrange
            $strEmail = "user@"

            # Act
            $boolResult = Test-IsValidEmail -Email $strEmail

            # Assert
            $boolResult | Should -BeFalse
        }
    }
}

# ==============================================================================
# Testing with Mocks
# ==============================================================================
# Mocking allows you to isolate the function under test from external dependencies.
# Use Mock to replace cmdlets like Get-Content, Invoke-RestMethod, etc.

Describe "Get-ProcessedData" {
    Context "When source file exists" {
        BeforeAll {
            # Mock Get-Content to return test data without accessing the file system
            Mock Get-Content {
                return @("Line 1", "Line 2", "Line 3")
            }
        }

        It "Returns success code 0" {
            # Arrange
            $arrResult = $null
            $strSource = "C:\test\data.txt"

            # Act
            $intReturnCode = Get-ProcessedData -ReferenceToResult ([ref]$arrResult) -Source $strSource

            # Assert
            $intReturnCode | Should -Be 0
        }

        It "Populates reference with file content" {
            # Arrange
            $arrResult = $null
            $strSource = "C:\test\data.txt"

            # Act
            [void](Get-ProcessedData -ReferenceToResult ([ref]$arrResult) -Source $strSource)

            # Assert
            $arrResult | Should -HaveCount 3
            $arrResult[0] | Should -Be "Line 1"
        }

        It "Calls Get-Content with the correct path" {
            # Arrange
            $arrResult = $null
            $strSource = "C:\test\specific.txt"

            # Act
            [void](Get-ProcessedData -ReferenceToResult ([ref]$arrResult) -Source $strSource)

            # Assert - Verify the mock was called with expected parameters
            Should -Invoke Get-Content -Times 1 -ParameterFilter {
                $Path -eq "C:\test\specific.txt"
            }
        }
    }

    Context "When source file does not exist" {
        BeforeAll {
            # Mock Get-Content to return $null (simulating file not found)
            Mock Get-Content {
                return $null
            }
        }

        It "Returns failure code -1" {
            # Arrange
            $arrResult = $null
            $strSource = "C:\nonexistent\file.txt"

            # Act
            $intReturnCode = Get-ProcessedData -ReferenceToResult ([ref]$arrResult) -Source $strSource

            # Assert
            $intReturnCode | Should -Be -1
        }
    }
}

# ==============================================================================
# Additional Assertion Examples
# ==============================================================================
# Pester provides many assertion operators. Here are common examples:

Describe "Assertion Examples" {
    Context "Common Should operators" {
        It "Should -Be for exact equality" {
            "Hello" | Should -Be "Hello"
            5 | Should -Be 5
        }

        It "Should -BeExactly for case-sensitive string comparison" {
            "Hello" | Should -BeExactly "Hello"
            # "Hello" | Should -BeExactly "hello"  # This would fail
        }

        It "Should -BeNullOrEmpty for null or empty values" {
            $null | Should -BeNullOrEmpty
            "" | Should -BeNullOrEmpty
            @() | Should -BeNullOrEmpty
        }

        It "Should -Not -BeNullOrEmpty for non-empty values" {
            "text" | Should -Not -BeNullOrEmpty
            @(1, 2, 3) | Should -Not -BeNullOrEmpty
        }

        It "Should -HaveCount for collection length" {
            @(1, 2, 3) | Should -HaveCount 3
        }

        It "Should -Contain for collection membership" {
            @("a", "b", "c") | Should -Contain "b"
        }

        It "Should -BeGreaterThan and -BeLessThan for comparisons" {
            10 | Should -BeGreaterThan 5
            5 | Should -BeLessThan 10
        }

        It "Should -Match for regex matching" {
            "Hello World" | Should -Match "World"
            "test@example.com" | Should -Match "@"
        }

        It "Should -Throw for exception testing" {
            { throw "Error!" } | Should -Throw
            { throw "Specific error" } | Should -Throw -ExpectedMessage "Specific error"
        }
    }
}
