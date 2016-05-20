using Microsoft.IdentityModel;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IdentityModel.Selectors;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.ServiceModel;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Script.Serialization;
using AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction;
using AudienceUriValidationFailedException = Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException;
using SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration;
using X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials;

namespace SharePoint.ReciveMailAddinWeb
{
    public static class TokenHelper
    {
        #region パブリック フィールド

        /// <summary>
        /// SharePoint プリンシパル。
        /// </summary>
        public const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// HighTrust アクセス トークンの有効期間 (12 時間)。
        /// </summary>
        public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        #endregion public fields

        #region パブリック メソッド

        /// <summary>
        /// POSTed フォーム パラメーターとクエリ文字列で既知のパラメーター名を探すことにより、
        /// 指定された要求からコンテキスト トークン文字列を取得します。コンテキスト トークンが見つかった場合は、null を返します。
        /// </summary>
        /// <param name="request">コンテキスト トークンを探す HttpRequest</param>
        /// <returns>コンテキスト トークン文字列</returns>
        public static string GetContextTokenFromRequest(HttpRequest request)
        {
            return GetContextTokenFromRequest(new HttpRequestWrapper(request));
        }

        /// <summary>
        /// POSTed フォーム パラメーターとクエリ文字列で既知のパラメーター名を探すことにより、
        /// 指定された要求からコンテキスト トークン文字列を取得します。コンテキスト トークンが見つからない場合は、null を返します。
        /// </summary>
        /// <param name="request">コンテキスト トークンを探す HttpRequest</param>
        /// <returns>コンテキスト トークン文字列</returns>
        public static string GetContextTokenFromRequest(HttpRequestBase request)
        {
            string[] paramNames = { "AppContext", "AppContextToken", "AccessToken", "SPAppToken" };
            foreach (string paramName in paramNames)
            {
                if (!string.IsNullOrEmpty(request.Form[paramName]))
                {
                    return request.Form[paramName];
                }
                if (!string.IsNullOrEmpty(request.QueryString[paramName]))
                {
                    return request.QueryString[paramName];
                }
            }
            return null;
        }

        /// <summary>
        /// web.config で指定されたパラメーターに基づいて、指定されたコンテキスト トークン文字列がこのアプリケーション用で 
        /// あることを検証します。検証に使用される web.config から使用されるパラメーターには ClientId が含まれます。
        /// HostedAppHostNameOverride、HostedAppHostName、ClientSecret、およびレルム (指定されている場合) が含まれます。HostedAppHostNameOverride が存在する場合は、
        ///検証に使用されます。それ以外では、<paramref name="appHostName"/> が 
        /// null でない場合は、web.config の HostedAppHostName の代わりに検証に使用されます。トークンが無効な場合、
        /// 例外がスローされます。トークンが有効な場合、TokenHelper の静的 STS メタデータ URL はトークンの内容に基づいて更新され、
        /// コンテキスト トークンに基づく JsonWebSecurityToken が返されます。
        /// </summary>
        /// <param name="contextTokenString">検証するコンテキスト トークン</param>
        /// <param name="appHostName">トークン対象ユーザーの検証に使用する URL 機関。ドメイン ネーム システム (DNS) ホスト名または IP アドレスとポート番号で構成されます。
        ///null の場合は、代わりに HostedAppHostName web.config 設定が使用されます。HostedAppHostNameOverride web.config 設定 (存在する場合) は 
        ///<paramref name="appHostName"/> の代わりに検証に使用されます。</param>
        /// <returns>コンテキスト トークンに基づく JsonWebSecurityToken。</returns>
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            JsonWebSecurityTokenHandler tokenHandler = CreateJsonWebSecurityTokenHandler();
            SecurityToken securityToken = tokenHandler.ReadToken(contextTokenString);
            JsonWebSecurityToken jsonToken = securityToken as JsonWebSecurityToken;
            SharePointContextToken token = SharePointContextToken.Create(jsonToken);

            string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            int firstDot = stsAuthority.IndexOf('.');

            GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            tokenHandler.ValidateToken(jsonToken);

            string[] acceptableAudiences;
            if (!String.IsNullOrEmpty(HostedAppHostNameOverride))
            {
                acceptableAudiences = HostedAppHostNameOverride.Split(';');
            }
            else if (appHostName == null)
            {
                acceptableAudiences = new[] { HostedAppHostName };
            }
            else
            {
                acceptableAudiences = new[] { appHostName };
            }

            bool validationSuccessful = false;
            string realm = Realm ?? token.Realm;
            foreach (var audience in acceptableAudiences)
            {
                string principal = GetFormattedPrincipal(ClientId, audience, realm);
                if (StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal))
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
            {
                throw new AudienceUriValidationFailedException(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", String.Join(";", acceptableAudiences), token.Audience));
            }

            return token;
        }

        /// <summary>
        /// ACS からアクセス トークンを取得し、指定された targetHost で指定されたコンテキスト トークンのソースを
        /// 呼び出します。targetHost は、コンテキスト トークンを送信したプリンシパルに登録されている必要があります。
        /// </summary>
        /// <param name="contextToken">意図されたアクセス トークン対象ユーザーにより発行されたコンテキスト</param>
        /// <param name="targetHost">ターゲット プリンシパルの URL 機関</param>
        /// <returns>対象ユーザーがコンテキスト トークンのソースと一致するアクセス トークン</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(SharePointContextToken contextToken, string targetHost)
        {
            string targetPrincipalName = contextToken.TargetPrincipalName;

            // コンテキスト トークンから refreshToken を抽出します
            string refreshToken = contextToken.RefreshToken;

            if (String.IsNullOrEmpty(refreshToken))
            {
                return null;
            }

            string targetRealm = Realm ?? contextToken.Realm;

            return GetAccessToken(refreshToken,
                                  targetPrincipalName,
                                  targetHost,
                                  targetRealm);
        }

        /// <summary>
        /// 指定された認証コードを使用し、ACS からアクセス トークンを取得して指定された targetHost で指定されたプリンシパルを
        /// 呼び出します。targetHost は、ターゲット プリンシパルに登録されている必要があります。指定されたレルムが
        /// null の場合、web.config の "Realm" 設定が代わりに使用されます。
        /// </summary>
        /// <param name="authorizationCode">アクセス トークンを交換するための認証コード</param>
        /// <param name="targetPrincipalName">アクセス トークンを取得するターゲット プリンシパルの名前</param>
        /// <param name="targetHost">ターゲット プリンシパルの URL 機関</param>
        /// <param name="targetRealm">アクセス トークンの nameid と対象ユーザーに使用するレルム</param>
        /// <param name="redirectUri">このアプリ用に登録されているリダイレクト URI</param>
        /// <returns>ターゲット プリンシパルの対象ユーザーを持つアクセス トークン</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(
            string authorizationCode,
            string targetPrincipalName,
            string targetHost,
            string targetRealm,
            Uri redirectUri)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

            // トークンの要求を作成します。ここでは、RedirectUri は null です。リダイレクト URI が登録されている場合は失敗します
            OAuth2AccessTokenRequest oauth2Request =
                OAuth2MessageFactory.CreateAccessTokenRequestWithAuthorizationCode(
                    clientId,
                    ClientSecret,
                    authorizationCode,
                    redirectUri,
                    resource);

            // トークンを取得します
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// 指定された更新トークンを使用して、ACS からアクセス トークンを取得し、指定された targetHost で指定されたプリンシパルを
        /// 呼び出します。targetHost は、ターゲット プリンシパルに登録されている必要があります。指定されたレルムが
        /// null の場合、web.config の "Realm" 設定が代わりに使用されます。
        /// </summary>
        /// <param name="refreshToken">アクセス トークンを交換するためのトークンの更新</param>
        /// <param name="targetPrincipalName">アクセス トークンを取得するターゲット プリンシパルの名前</param>
        /// <param name="targetHost">ターゲット プリンシパルの URL 機関</param>
        /// <param name="targetRealm">アクセス トークンの nameid と対象ユーザーに使用するレルム</param>
        /// <returns>ターゲット プリンシパルの対象ユーザーを持つアクセス トークン</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(
            string refreshToken,
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(clientId, ClientSecret, refreshToken, resource);

            // トークンを取得します
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// ACS からアプリケーション専用のアクセス トークンを取得し、指定されたプリンシパルを呼び出します
        /// 呼び出します。targetHost は、ターゲット プリンシパルに登録されている必要があります。指定されたレルムが
        /// null の場合、web.config の "Realm" 設定が代わりに使用されます。
        /// </summary>
        /// <param name="targetPrincipalName">アクセス トークンを取得するターゲット プリンシパルの名前</param>
        /// <param name="targetHost">ターゲット プリンシパルの URL 機関</param>
        /// <param name="targetRealm">アクセス トークンの nameid と対象ユーザーに使用するレルム</param>
        /// <returns>ターゲット プリンシパルの対象ユーザーを持つアクセス トークン</returns>
        public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {

            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, ClientSecret, resource);
            oauth2Request.Resource = resource;

            // トークンを取得します
            OAuth2S2SClient client = new OAuth2S2SClient();

            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// リモート イベント レシーバーのプロパティに基づいてクライアント コンテキストを作成します
        /// </summary>
        /// <param name="properties">リモート イベント レシーバーのプロパティ</param>
        /// <returns>イベントが発生した Web の呼び出し準備ができた ClientContext</returns>
        public static ClientContext CreateRemoteEventReceiverClientContext(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl;
            if (properties.ListEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ListEventProperties.WebUrl);
            }
            else if (properties.ItemEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            }
            else if (properties.WebEventProperties != null)
            {
                sharepointUrl = new Uri(properties.WebEventProperties.FullUrl);
            }
            else
            {
                return null;
            }

            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// アプリ イベントのプロパティに基づいてクライアント コンテキストを作成します
        /// </summary>
        /// <param name="properties">アプリ イベントのプロパティ</param>
        /// <param name="useAppWeb">アプリ Web を対象にする場合は True、ホスト Web を対象にする場合は False</param>
        /// <returns>アプリ Web または親 Web の呼び出し準備ができた ClientContext</returns>
        public static ClientContext CreateAppEventClientContext(SPRemoteEventProperties properties, bool useAppWeb)
        {
            if (properties.AppEventProperties == null)
            {
                return null;
            }

            Uri sharepointUrl = useAppWeb ? properties.AppEventProperties.AppWebFullUrl : properties.AppEventProperties.HostWebFullUrl;
            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// 指定された認証コードを使用して ACS からアクセス トークンを取得し、そのアクセス トークンを使用して
        /// クライアント コンテキストを作成します
        /// </summary>
        /// <param name="targetUrl">ターゲット SharePoint サイトの URL</param>
        /// <param name="authorizationCode">ACS からアクセス トークンを取得するときに使用する認証コード</param>
        /// <param name="redirectUri">このアプリ用に登録されているリダイレクト URI</param>
        /// <returns>有効なアクセス トークンを持つ targetUrl の呼び出し準備ができた ClientContext</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string authorizationCode,
            Uri redirectUri)
        {
            return GetClientContextWithAuthorizationCode(targetUrl, SharePointPrincipal, authorizationCode, GetRealmFromTargetUrl(new Uri(targetUrl)), redirectUri);
        }

        /// <summary>
        /// 指定された認証コードを使用して ACS からアクセス トークンを取得し、そのアクセス トークンを使用して
        /// クライアント コンテキストを作成します
        /// </summary>
        /// <param name="targetUrl">ターゲット SharePoint サイトの URL</param>
        /// <param name="targetPrincipalName">ターゲット SharePoint プリンシパルの名前</param>
        /// <param name="authorizationCode">ACS からアクセス トークンを取得するときに使用する認証コード</param>
        /// <param name="targetRealm">アクセス トークンの nameid と対象ユーザーに使用するレルム</param>
        /// <param name="redirectUri">このアプリ用に登録されているリダイレクト URI</param>
        /// <returns>有効なアクセス トークンを持つ targetUrl の呼び出し準備ができた ClientContext</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string targetPrincipalName,
            string authorizationCode,
            string targetRealm,
            Uri redirectUri)
        {
            Uri targetUri = new Uri(targetUrl);

            string accessToken =
                GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// 指定されたアクセス トークンを使用してクライアント コンテキストを作成します
        /// </summary>
        /// <param name="targetUrl">ターゲット SharePoint サイトの URL</param>
        /// <param name="accessToken">指定された targetUrl を呼び出すときに使用するアクセス トークン</param>
        /// <returns>指定されたアクセス トークンを持つ targetUrl の呼び出し準備ができた ClientContext</returns>
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);

            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate(object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        /// <summary>
        /// 指定されたコンテキスト トークンを使用して ACS からアクセス トークンを取得し、そのアクセス トークンを使用して
        /// クライアント コンテキストを作成します
        /// </summary>
        /// <param name="targetUrl">ターゲット SharePoint サイトの URL</param>
        /// <param name="contextTokenString">ターゲット SharePoint サイトから取得したトークン</param>
        /// <param name="appHostUrl">ホストされるアプリの URL 機関。これが null の場合、
        /// HostedAppHostName の値が代わりに使用されます</param>
        /// <returns>有効なアクセス トークンを持つ targetUrl の呼び出し準備ができた ClientContext</returns>
        public static ClientContext GetClientContextWithContextToken(
            string targetUrl,
            string contextTokenString,
            string appHostUrl)
        {
            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, appHostUrl);

            Uri targetUri = new Uri(targetUrl);

            string accessToken = GetAccessToken(contextToken, targetUri.Authority).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// アプリケーションがブラウザーをリダイレクトして同意を求める SharePoint URL を返し、
        /// 認証コードを返します。
        /// </summary>
        /// <param name="contextUrl">SharePoint サイトの絶対 URL</param>
        /// <param name="scope">SharePoint サイトから要求する "短縮形" のスペース区切りのアクセス許可 
        /// ( "Web.Read Site.Write" など)</param>
        /// <returns>SharePoint サイトの OAuth 認証ページの URL</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope);
        }

        /// <summary>
        /// アプリケーションがブラウザーをリダイレクトして同意を求める SharePoint URL を返し、
        /// 認証コードを返します。
        /// </summary>
        /// <param name="contextUrl">SharePoint サイトの絶対 URL</param>
        /// <param name="scope">SharePoint サイトから要求する "短縮形" のスペース区切りのアクセス許可
        /// ( "Web.Read Site.Write" など)</param>
        /// <param name="redirectUri">同意を得た後に SharePoint がブラウザーをリダイレクト
        /// する URI</param>
        /// <returns>SharePoint サイトの OAuth 認証ページの URL</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope, string redirectUri)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope,
                redirectUri);
        }

        /// <summary>
        /// アプリケーションがブラウザーをリダイレクトして新しいコンテキスト トークンを要求する SharePoint URL です。
        /// </summary>
        /// <param name="contextUrl">SharePoint サイトの絶対 URL</param>
        /// <param name="redirectUri">SharePoint がコンテキスト トークを使用してブラウザーをリダイレクトする URI</param>
        /// <returns>SharePoint サイトのコンテキスト トークン リダイレクト ページの URL</returns>
        public static string GetAppContextTokenRequestUrl(string contextUrl, string redirectUri)
        {
            return string.Format(
                "{0}{1}?client_id={2}&redirect_uri={3}",
                EnsureTrailingSlash(contextUrl),
                RedirectPage,
                ClientId,
                redirectUri);
        }

        /// <summary>
        ///  指定された WindowsIdentity の代わりにアプリケーションのプライベート証明書により署名された、
        ///targetApplicationUri の SharePoint 用の S2S アクセス トークンを取得します。web.config でレルムが指定されていない
        /// 場合、検出するために認証チャレンジが targetApplicationUri に発行されます。
        /// </summary>
        /// <param name="targetApplicationUri">ターゲット SharePoint サイトの URL</param>
        /// <param name="identity">代わりにアクセス トークンを作成するユーザーの Windows ID</param>
        /// <returns>ターゲット プリンシパルの対象ユーザーを持つアクセス トークン</returns>
        public static string GetS2SAccessTokenWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
        }

        /// <summary>
        /// 指定された WindowsIdentity の代わりにアプリケーションのプライベート証明書により署名された、
        /// targetApplicationUri のアプリケーション用のアクセス トークンを持つ S2S クライアント コンテキストを、targetRealm を使用して
        /// 取得します。web.config でレルムが指定されていない場合、検出するために認証チャレンジが 
        /// targetApplicationUri に発行されます。
        /// </summary>
        /// <param name="targetApplicationUri">ターゲット SharePoint サイトの URL</param>
        /// <param name="identity">代わりにアクセス トークンを作成するユーザーの Windows ID</param>
        /// <returns>ターゲット アプリケーションの対象ユーザーを持つアクセス トークンを使用した ClientContext</returns>
        public static ClientContext GetS2SClientContextWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            string accessToken = GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);

            return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }

        /// <summary>
        /// SharePoint の認証レルムの取得
        /// </summary>
        /// <param name="targetApplicationUri">ターゲット SharePoint サイトの URL</param>
        /// <returns>レルム GUID の文字列表現</returns>
        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 信頼性の高いアプリであるかどうかを判定します。
        /// </summary>
        /// <returns>信頼性の高いアプリの場合には True。</returns>
        public static bool IsHighTrustApp()
        {
            return SigningCredentials != null;
        }

        /// <summary>
        /// 指定された URL が null および空でない場合には、'/' で終わるようにします。
        /// </summary>
        /// <param name="url">URL。</param>
        /// <returns>null および空でない場合に '/' で終わる URL。</returns>
        public static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
            {
                return url + "/";
            }

            return url;
        }

        #endregion

        #region プライベート フィールド

        //
        // 構成定数
        //        

        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string S2SProtocol = "OAuth2";
        private const string DelegationIssuance = "DelegationIssuance1.0";
        private const string NameIdentifierClaimType = JsonWebTokenConstants.ReservedClaims.NameIdentifier;
        private const string TrustedForImpersonationClaimType = "trustedfordelegation";
        private const string ActorTokenClaimType = JsonWebTokenConstants.ReservedClaims.ActorToken;

        //
        // 環境定数
        //

        private static string GlobalEndPointPrefix = "accounts";
        private static string AcsHostUrl = "accesscontrol.windows.net";

        //
        // ホストされるアプリケーションの構成
        //
        private static readonly string ClientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId")) ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");
        private static readonly string IssuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");
        private static readonly string HostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
        private static readonly string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
        private static readonly string ClientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret")) ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");
        private static readonly string SecondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");
        private static readonly string Realm = WebConfigurationManager.AppSettings.Get("Realm");
        private static readonly string ServiceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string ClientSigningCertificatePath = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");
        private static readonly string ClientSigningCertificatePassword = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");
        private static readonly X509Certificate2 ClientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword)) ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
        private static readonly X509SigningCredentials SigningCredentials = (ClientCertificate == null) ? null : new X509SigningCredentials(ClientCertificate, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest);

        #endregion

        #region プライベート メソッド

        private static ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl)
        {
            string contextTokenString = properties.ContextToken;

            if (String.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }

            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
            string accessToken = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

            return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
        }

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return String.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
        }

        private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler()
        {
            JsonWebSecurityTokenHandler handler = new JsonWebSecurityTokenHandler();
            handler.Configuration = new SecurityTokenHandlerConfiguration();
            handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
            handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            List<byte[]> securityKeys = new List<byte[]>();
            securityKeys.Add(Convert.FromBase64String(ClientSecret));
            if (!string.IsNullOrEmpty(SecondaryClientSecret))
            {
                securityKeys.Add(Convert.FromBase64String(SecondaryClientSecret));
            }

            List<SecurityToken> securityTokens = new List<SecurityToken>();
            securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            handler.Configuration.IssuerTokenResolver =
                SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                new ReadOnlyCollection<SecurityToken>(securityTokens),
                false);
            SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            foreach (byte[] securitykey in securityKeys)
            {
                issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(ServiceNamespace));
            }
            handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        private static string GetS2SAccessTokenWithClaims(
            string targetApplicationHostName,
            string targetRealm,
            IEnumerable<JsonWebTokenClaim> claims)
        {
            return IssueToken(
                ClientId,
                IssuerId,
                targetRealm,
                SharePointPrincipal,
                targetRealm,
                targetApplicationHostName,
                true,
                claims,
                claims == null);
        }

        private static JsonWebTokenClaim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity)
        {
            JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
            {
                new JsonWebTokenClaim(NameIdentifierClaimType, identity.User.Value.ToLower()),
                new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
            };
            return claims;
        }

        private static string IssueToken(
            string sourceApplication,
            string issuerApplication,
            string sourceRealm,
            string targetApplication,
            string targetRealm,
            string targetApplicationHostName,
            bool trustedForDelegation,
            IEnumerable<JsonWebTokenClaim> claims,
            bool appOnly = false)
        {
            if (null == SigningCredentials)
            {
                throw new InvalidOperationException("SigningCredentials was not initialized");
            }

            #region アクター トークン

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<JsonWebTokenClaim> actorClaims = new List<JsonWebTokenClaim>();
            actorClaims.Add(new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid));
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new JsonWebTokenClaim(TrustedForImpersonationClaimType, "true"));
            }

            // トークンを作成します
            JsonWebSecurityToken actorToken = new JsonWebSecurityToken(
                issuer: issuer,
                audience: audience,
                validFrom: DateTime.UtcNow,
                validTo: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                signingCredentials: SigningCredentials,
                claims: actorClaims);

            string actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);

            if (appOnly)
            {
                // アプリケーション専用トークンはが、委任されたケースのアクター トークンと同じです
                return actorTokenString;
            }

            #endregion Actor token

            #region 外側トークン

            List<JsonWebTokenClaim> outerClaims = null == claims ? new List<JsonWebTokenClaim>() : new List<JsonWebTokenClaim>(claims);
            outerClaims.Add(new JsonWebTokenClaim(ActorTokenClaimType, actorTokenString));

            JsonWebSecurityToken jsonToken = new JsonWebSecurityToken(
                nameid, // 外部トークンの発行者はアクター トークンの nameid と一致する必要があります
                audience,
                DateTime.UtcNow,
                DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                outerClaims);

            string accessToken = new JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken);

            #endregion Outer token

            return accessToken;
        }

        #endregion

        #region AcsMetadataParser

        // このクラスは、グローバル STS エンドポイントから MetaData ドキュメントを取得するために使用されます。これには、
        // MetaData ドキュメントを解析し、エンドポイントと STS 証明書を取得するメソッドが含まれています。
        public static class AcsMetadataParser
        {
            public static X509Certificate2 GetAcsSigningCert(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                if (null != document.keys && document.keys.Count > 0)
                {
                    JsonKey signingKey = document.keys[0];

                    if (null != signingKey && null != signingKey.keyValue)
                    {
                        return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                    }
                }

                throw new Exception("Metadata document does not contain ACS signing certificate.");
            }

            public static string GetDelegationServiceUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);

                if (null != delegationEndpoint)
                {
                    return delegationEndpoint.location;
                }
                throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
            }

            private static JsonMetadataDocument GetMetadataDocument(string realm)
            {
                string acsMetadataEndpointUrlWithRealm = String.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                       GetAcsMetadataEndpointUrl(),
                                                                       realm);
                byte[] acsMetadata;
                using (WebClient webClient = new WebClient())
                {

                    acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
                }
                string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

                if (null == document)
                {
                    throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
                }

                return document;
            }

            public static string GetStsUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

                if (null != s2sEndpoint)
                {
                    return s2sEndpoint.location;
                }

                throw new Exception("Metadata document does not contain STS endpoint url");
            }

            private class JsonMetadataDocument
            {
                public string serviceName { get; set; }
                public List<JsonEndpoint> endpoints { get; set; }
                public List<JsonKey> keys { get; set; }
            }

            private class JsonEndpoint
            {
                public string location { get; set; }
                public string protocol { get; set; }
                public string usage { get; set; }
            }

            private class JsonKeyValue
            {
                public string type { get; set; }
                public string value { get; set; }
            }

            private class JsonKey
            {
                public string usage { get; set; }
                public JsonKeyValue keyValue { get; set; }
            }
        }

        #endregion
    }

    /// <summary>
    /// サード パーティ アプリケーションを認証し、更新トークンを使用したコールバックを許可するために SharePoint により生成される JsonWebSecurityToken
    /// </summary>
    public class SharePointContextToken : JsonWebSecurityToken
    {
        public static SharePointContextToken Create(JsonWebSecurityToken contextToken)
        {
            return new SharePointContextToken(contextToken.Issuer, contextToken.Audience, contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims);
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims)
            : base(issuer, audience, validFrom, validTo, claims)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SecurityToken issuerToken, JsonWebSecurityToken actorToken)
            : base(issuer, audience, validFrom, validTo, claims, issuerToken, actorToken)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SigningCredentials signingCredentials)
            : base(issuer, audience, validFrom, validTo, claims, signingCredentials)
        {
        }

        public string NameId
        {
            get
            {
                return GetClaimValue(this, "nameid");
            }
        }

        /// <summary>
        /// コンテキスト トークンの "appctxsender" クレームのプリンシパル名部分
        /// </summary>
        public string TargetPrincipalName
        {
            get
            {
                string appctxsender = GetClaimValue(this, "appctxsender");

                if (appctxsender == null)
                {
                    return null;
                }

                return appctxsender.Split('@')[0];
            }
        }

        /// <summary>
        /// コンテキスト トークンの "refreshtoken" クレーム
        /// </summary>
        public string RefreshToken
        {
            get
            {
                return GetClaimValue(this, "refreshtoken");
            }
        }

        /// <summary>
        /// コンテキスト トークンの "CacheKey" クレーム
        /// </summary>
        public string CacheKey
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string cacheKey = (string)dict["CacheKey"];

                return cacheKey;
            }
        }

        /// <summary>
        /// コンテキスト トークンの "SecurityTokenServiceUri" クレーム
        /// </summary>
        public string SecurityTokenServiceUri
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string securityTokenServiceUri = (string)dict["SecurityTokenServiceUri"];

                return securityTokenServiceUri;
            }
        }

        /// <summary>
        /// コンテキスト トークンの "audience" クレームのレルム部分
        /// </summary>
        public string Realm
        {
            get
            {
                string aud = Audience;
                if (aud == null)
                {
                    return null;
                }

                string tokenRealm = aud.Substring(aud.IndexOf('@') + 1);

                return tokenRealm;
            }
        }

        private static string GetClaimValue(JsonWebSecurityToken token, string claimType)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }

            foreach (JsonWebTokenClaim claim in token.Claims)
            {
                if (StringComparer.Ordinal.Equals(claim.ClaimType, claimType))
                {
                    return claim.Value;
                }
            }

            return null;
        }

    }

    /// <summary>
    /// 対称アルゴリズムを使用して生成される複数のセキュリティ キーを含むセキュリティ トークンを表します。
    /// </summary>
    public class MultipleSymmetricKeySecurityToken : SecurityToken
    {
        /// <summary>
        /// MultipleSymmetricKeySecurityToken クラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="keys">対称キーを含むバイト配列の列挙。</param>
        public MultipleSymmetricKeySecurityToken(IEnumerable<byte[]> keys)
            : this(UniqueId.CreateUniqueId(), keys)
        {
        }

        /// <summary>
        /// MultipleSymmetricKeySecurityToken クラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="tokenId">セキュリティ トークンの一意識別子。</param>
        /// <param name="keys">対称キーを含むバイト配列の列挙。</param>
        public MultipleSymmetricKeySecurityToken(string tokenId, IEnumerable<byte[]> keys)
        {
            if (keys == null)
            {
                throw new ArgumentNullException("keys");
            }

            if (String.IsNullOrEmpty(tokenId))
            {
                throw new ArgumentException("Value cannot be a null or empty string.", "tokenId");
            }

            foreach (byte[] key in keys)
            {
                if (key.Length <= 0)
                {
                    throw new ArgumentException("The key length must be greater then zero.", "keys");
                }
            }

            id = tokenId;
            effectiveTime = DateTime.UtcNow;
            securityKeys = CreateSymmetricSecurityKeys(keys);
        }

        /// <summary>
        /// セキュリティ トークンの一意識別子を取得します。
        /// </summary>
        public override string Id
        {
            get
            {
                return id;
            }
        }

        /// <summary>
        /// セキュリティ トークンに関連付けられている暗号化キーを取得します。
        /// </summary>
        public override ReadOnlyCollection<SecurityKey> SecurityKeys
        {
            get
            {
                return securityKeys.AsReadOnly();
            }
        }

        /// <summary>
        /// このセキュリティ トークンが有効である最初の時点を取得します。
        /// </summary>
        public override DateTime ValidFrom
        {
            get
            {
                return effectiveTime;
            }
        }

        /// <summary>
        /// このセキュリティ トークンが有効である最後の時点を取得します。
        /// </summary>
        public override DateTime ValidTo
        {
            get
            {
                // 期限なし
                return DateTime.MaxValue;
            }
        }

        /// <summary>
        /// このインスタンスのキー識別子を、指定されたキー識別子に解決できるかどうかを示す値を返します。
        /// </summary>
        /// <param name="keyIdentifierClause">このインスタンスと比較する SecurityKeyIdentifierClause</param>
        /// <returns>keyIdentifierClause が SecurityKeyIdentifierClause で、一意識別子が ID プロパティと同じである場合は true、それ以外の場合は false。</returns>
        public override bool MatchesKeyIdentifierClause(SecurityKeyIdentifierClause keyIdentifierClause)
        {
            if (keyIdentifierClause == null)
            {
                throw new ArgumentNullException("keyIdentifierClause");
            }

            // これは対称トークンで、トークンを区別する ID はないため、
            // SymmetricIssuerKeyIdentifier があるかどうかをチェックするだけです。キーが発行者に一致する場合、発行者への実際の
            // マッピングは後で行われます。
            if (keyIdentifierClause is SymmetricIssuerKeyIdentifierClause)
            {
                return true;
            }
            return base.MatchesKeyIdentifierClause(keyIdentifierClause);
        }

        #region プライベート メンバー

        private List<SecurityKey> CreateSymmetricSecurityKeys(IEnumerable<byte[]> keys)
        {
            List<SecurityKey> symmetricKeys = new List<SecurityKey>();
            foreach (byte[] key in keys)
            {
                symmetricKeys.Add(new InMemorySymmetricSecurityKey(key));
            }
            return symmetricKeys;
        }

        private string id;
        private DateTime effectiveTime;
        private List<SecurityKey> securityKeys;

        #endregion
    }
}
