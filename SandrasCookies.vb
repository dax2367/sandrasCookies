Option Strict On

Public Class SandrasCookies
    '=============================================================================================================
    ' Author:	Holly Eaton
    ' 
    ' Program:	Sandra's Cookies
    ' 
    ' Dev Env:	Visual Studio
    ' 
    ' Description:
    '   Purpose:    Project that will determine how much a customer owes for an order of Sandra's Cookies.
    '
    '   Input:      The number of large and small cookies.
    '
    '  	Process:    Calculate the subtotals for cookies, sales tax and shipping and the Grand total for the order.
    '
    ' 	Output:     Textual information for the user inside the labels(totals) and textboxes(number of cookies).
    ' 
    '===============================================================================================================
    ' 	Declared Constants:
    '	SngLG_CK_PRICE
    '	SngREG_CK_PRICE
    '	SngLG_CK_OUNCES
    '	SngREG_CK_OUNCES
    '	SngSALES_TAX_RATE
    '	SngSHIP_CHARGE_PER_OUNCE
    '
    '===============================================================================================================
    '	Variables for user entered data:
    '	SngNumLgCookies
    '	SngNumRegCookies
    '	Note: (dim and cast both directly to single then no type casting necessary in calculations)
    '	Example: Dim sngNumRegCookies As Single = Csng(txtNumRegCookies.Text)
    '
    '===============================================================================================================
    '	Variables for calculated values:
    '	SngCookieCost
    '	SngSalesTaxCost
    '	SngTotalOz
    '	SngShippingCost
    '	SngTotalCost
    '
    '================================================================================================================
    '	Calculations in pseudocode
    '	Set option strict on at top (since everything is sng won’t matter in calculations
    '   but will when casting sng to string in label output)
    '
    '	CookieCost = (NumLgCookies * Lg_Ck_Price) + (NumRegCookies * Reg_Ck_Price)
    '	SalesTaxCost = CookieCost * SALES_TAX_RATE
    '	TotalOz = (NumLgCookies * Lg_Ck_Ounces) + (NumRegCookies * Reg_Ck_ Ounces)
    '	ShippingCost = (TotalOz * Ship_Charge_Per_Ounce)
    '	TotalCost = CookieCost + SalesTaxCost  + ShippingCost 
    '
    '==================================================================================================================
    '==================================================================================================================
    '==================================================================================================================

    'Declared Constants:
    Const SngLG_CK_PRICE As Single = 1.5
    Const SngREG_CK_PRICE As Single = 1.0
    Const SngLG_CK_OUNCES As Single = 7.0
    Const SngREG_CK_OUNCES As Single = 4.5
    Const SngSALES_TAX_RATE As Single = 0.0875
    Const SngSHIP_CHARGE_PER_OUNCE As Single = 0.042

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Variables for calculated values:
        Dim SngCookieCost As Single
        Dim SngSalesTaxCost As Single
        Dim SngTotalOz As Single
        Dim SngShippingCost As Single
        Dim SngTotalCost As Single

        'Variables for user entered data:   Note: (dim and cast both directly to single then no type casting necessary in calculations)
        Try
            Dim sngNumLgCookies As Single = CSng(txtNumLgCookies.Text)

            Try
                Dim sngNumRegCookies As Single = CSng(txtNumRegCookies.Text)

                'Calculations for the subtotals for cookies, sales tax and shipping and the Grand total for the order.
                '   Cookie Cost
                SngCookieCost = (sngNumLgCookies * SngLG_CK_PRICE) + (sngNumRegCookies * SngREG_CK_PRICE)
                '   Sales Tax Cost
                SngSalesTaxCost = SngCookieCost * SngSALES_TAX_RATE
                '   Total Ounces
                SngTotalOz = (sngNumLgCookies * SngLG_CK_OUNCES) + (sngNumRegCookies * SngREG_CK_OUNCES)
                '   Shipping Cost
                SngShippingCost = (SngTotalOz * SngSHIP_CHARGE_PER_OUNCE)
                '   Total Cost
                SngTotalCost = SngCookieCost + SngSalesTaxCost + SngShippingCost

                'format and output the cookie results
                lblCookieCost.Text = SngCookieCost.ToString("c")
                lblSalesTaxCost.Text = SngSalesTaxCost.ToString("c")
                lblShippingCost.Text = SngShippingCost.ToString("c")
                lblTotalCost.Text = SngTotalCost.ToString("c")
            Catch ex As Exception
                'What to do if user data entered into txtNumRegCookies is invalid and cannot be cast to single.
                'Tell user what to enter
                MessageBox.Show("Please enter the number of regular cookies you want using only numerical characters.                                                                                                                        For example, if you want a dozen cookies enter 12.                                                     Twelve or dozen is invalid. Thank you.")

                'reset input area
                txtNumRegCookies.Text = "0"

                'put insertion point inside of Lg cookie textbox
                txtNumRegCookies.Focus()
            End Try

        Catch ex As Exception
            'What to do if user data entered into txtNumLgCookies is invalid and cannot be cast to single.
            'Tell user what to enter
            MessageBox.Show("Please enter the number of large cookies you want using only numerical characters.                                                                                                                        For example, if you want a dozen cookies enter 12.                                                     Twelve or dozen is invalid. Thank you.")

            'reset input area
            txtNumLgCookies.Text = "0"

            'put insertion point inside of Lg cookie textbox
            txtNumLgCookies.Focus()

        End Try

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear the textboxes and labels
        txtNumLgCookies.Text = "0"
        txtNumRegCookies.Text = "0"
        lblCookieCost.Text = String.Empty
        lblSalesTaxCost.Text = String.Empty
        lblShippingCost.Text = String.Empty
        lblTotalCost.Text = String.Empty

        'Give the focus to txtNumLgCookies
        txtNumLgCookies.Focus()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Close the form
        Me.Close()
    End Sub
End Class
