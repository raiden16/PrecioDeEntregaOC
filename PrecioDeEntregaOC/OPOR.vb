Public Class OPOR

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private Directorio As String

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Directorio = oCatchingEvents.csDirectory
    End Sub

    '//----- AGREGA ELEMENTOS A LA FORMA
    Public Sub addFormItems(ByVal FormUID As String)
        Dim loItem As SAPbouiCOM.Item
        Dim loButton As SAPbouiCOM.Button
        Dim lsItemRef As String

        Try

            '//AGREGA BOTON MOVIMIENTOS EN PEDIDOS DE COMPRAS
            coForm = cSBOApplication.Forms.Item(FormUID)

            'If coForm.Items.Item("btCost").Enabled = True Then
            'lsItemRef = "btCost"
            'Else
            lsItemRef = "2"
            'End If
            '"btPriDl" Para que funcione primero corre este, despues el de ExtraPriceItems y por ultimo el de movimientos

            loItem = coForm.Items.Add("btPriDl", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            If cSBOCompany.UserSignature = 25 Or cSBOCompany.UserSignature = 1 Then
                loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + coForm.Items.Item(lsItemRef).Width + coForm.Items.Item(lsItemRef).Width + coForm.Items.Item(lsItemRef).Width + 30
            ElseIf cSBOCompany.UserSignature = 37 Then
                loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 10
            End If
            loItem.Top = coForm.Items.Item(lsItemRef).Top
            loItem.Width = coForm.Items.Item(lsItemRef).Width + 40
            loItem.Height = coForm.Items.Item(lsItemRef).Height
            loButton = loItem.Specific
            loButton.Caption = "Precio de Entrega"

        Catch ex As Exception
            cSBOApplication.MessageBox("DocumentoSBO. agregar elementos a la forma. " & ex.Message)
        Finally
            coForm = Nothing
            loItem = Nothing
            loButton = Nothing
        End Try
    End Sub

End Class
