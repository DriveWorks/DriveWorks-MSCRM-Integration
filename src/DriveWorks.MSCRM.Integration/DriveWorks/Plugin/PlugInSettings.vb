Imports DriveWorks.Applications
Imports System.Security.Cryptography

''' <summary>
''' Provides access to the plugins settings.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class PluginSettings

    ' Registry address - DO NOT CHANGE.
    Private Const SETTING_BASE As String = "Common\MSCRM\Integration\"
    Private Const SETTING_USER_NAME As String = SETTING_BASE & "UserName"
    Private Const SETTING_PASSWORD As String = SETTING_BASE & "Password"
    Private Const SETTING_CONNECTION_STRING As String = SETTING_BASE & "ConnectionString"

    Private ReadOnly mSettingsManager As ISettingsManager

    Public Sub New(ByVal settingsManager As ISettingsManager)
        mSettingsManager = settingsManager
    End Sub

#Region " Properties "

    ''' <summary>
    ''' Gets/sets the username used to connect to the 3rd party system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Property UserName() As String
        Get
            Return mSettingsManager.GetSettingAsString(SettingLevel.User, SETTING_USER_NAME, False)
        End Get
        Set(ByVal value As String)
            mSettingsManager.SetSetting(SettingLevel.User, SETTING_USER_NAME, value, False)
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the password to use to connect to the 3rd party system.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Password() As String
        Get
            Return mSettingsManager.GetSettingAsDecryptedString(SettingLevel.User, SETTING_PASSWORD, False)
        End Get
        Set(ByVal value As String)
            mSettingsManager.SetSettingEncrypted(SettingLevel.User, SETTING_PASSWORD, value, False)
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the connection string used to connect to the 3rd party system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Property ConnectionString() As String
        Get
            Return mSettingsManager.GetSettingAsString(SettingLevel.User, SETTING_CONNECTION_STRING, False)
        End Get
        Set(ByVal value As String)
            mSettingsManager.SetSetting(SettingLevel.User, SETTING_CONNECTION_STRING, value, False)
        End Set
    End Property

#End Region

End Class