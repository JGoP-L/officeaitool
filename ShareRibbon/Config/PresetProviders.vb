Public Class PresetProviders
    Public Shared Function GetCloudProviders() As List(Of ConfigManager.ConfigItem)
        Return New List(Of ConfigManager.ConfigItem)()
    End Function

    Public Shared Function GetLocalProviders() As List(Of ConfigManager.ConfigItem)
        Return New List(Of ConfigManager.ConfigItem)()
    End Function
End Class
