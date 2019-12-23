// ------------------------------------------------------------------------------
//  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

namespace Microsoft.Graph
{
    using System;
    using System.Net.Http;
    using System.Security;
    using Microsoft.Identity.Client;

    /// <summary>
    /// Extension methods for <see cref="BaseRequest"/>
    /// </summary>
    public static class BaseRequestExtensions
    {

        /// <summary>
        /// Sets the default authentication provider to the default Authentication Middleware Handler for this request.
        /// This only works with the default authentication handler.
        /// If you use a custom authentication handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <returns></returns>
        internal static T WithDefaultAuthProvider<T>(this T baseRequest) where T : IBaseRequest
        {
            string authOptionKey = typeof(AuthenticationHandlerOption).ToString();
            if (baseRequest.MiddlewareOptions.ContainsKey(authOptionKey))
            {
                (baseRequest.MiddlewareOptions[authOptionKey] as AuthenticationHandlerOption).AuthenticationProvider = baseRequest.Client.AuthenticationProvider;
            }
            else
            {
                baseRequest.MiddlewareOptions.Add(authOptionKey, new AuthenticationHandlerOption { AuthenticationProvider = baseRequest.Client.AuthenticationProvider });
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets the PerRequestAuthProvider delegate handler to the default Authentication Middleware Handler to authenticate a single request.
        /// The PerRequestAuthProvider delegate handler must be set to the GraphServiceClient instance before using this extension method otherwise, it defaults to the default authentication provider.
        /// This only works with the default authentication handler.
        /// If you use a custom authentication handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <returns></returns>
        public static T WithPerRequestAuthProvider<T>(this T baseRequest) where T : IBaseRequest
        {
            if (baseRequest.Client.PerRequestAuthProvider != null)
            {
                string authOptionKey = typeof(AuthenticationHandlerOption).ToString();
                if (baseRequest.MiddlewareOptions.ContainsKey(authOptionKey))
                {
                    (baseRequest.MiddlewareOptions[authOptionKey] as AuthenticationHandlerOption).AuthenticationProvider = baseRequest.Client.PerRequestAuthProvider();
                }
                else
                {
                    baseRequest.MiddlewareOptions.Add(authOptionKey, new AuthenticationHandlerOption { AuthenticationProvider = baseRequest.Client.PerRequestAuthProvider() });
                }
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets a ShouldRetry delegate to the default Retry Middleware Handler for this request.
        /// This only works with the default Retry Middleware Handler.
        /// If you use a custom Retry Middleware Handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <param name="shouldRetry">A <see cref="Func{Int32, Int32, HttpResponseMessage, Boolean}"/> for the request.</param>
        /// <returns></returns>
        public static T WithShouldRetry<T>(this T baseRequest, Func<int, int, HttpResponseMessage, bool> shouldRetry) where T : IBaseRequest
        {
            string retryOptionKey = typeof(RetryHandlerOption).ToString();
            if (baseRequest.MiddlewareOptions.ContainsKey(retryOptionKey))
            {
                (baseRequest.MiddlewareOptions[retryOptionKey] as RetryHandlerOption).ShouldRetry = shouldRetry;
            }
            else
            {
                baseRequest.MiddlewareOptions.Add(retryOptionKey, new RetryHandlerOption { ShouldRetry = shouldRetry });
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets the maximum number of retries to the default Retry Middleware Handler for this request.
        /// This only works with the default Retry Middleware Handler.
        /// If you use a custom Retry Middleware Handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <param name="maxRetry">The maxRetry for the request.</param>
        /// <returns></returns>
        public static T WithMaxRetry<T>(this T baseRequest, int maxRetry) where T : IBaseRequest
        {
            string retryOptionKey = typeof(RetryHandlerOption).ToString();
            if (baseRequest.MiddlewareOptions.ContainsKey(retryOptionKey))
            {
                (baseRequest.MiddlewareOptions[retryOptionKey] as RetryHandlerOption).MaxRetry = maxRetry;
            }
            else
            {
                baseRequest.MiddlewareOptions.Add(retryOptionKey, new RetryHandlerOption { MaxRetry = maxRetry });
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets the maximum time for request retries to the default Retry Middleware Handler for this request.
        /// This only works with the default Retry Middleware Handler.
        /// If you use a custom Retry Middleware Handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <param name="retriesTimeLimit">The retriestimelimit for the request in seconds.</param>
        /// <returns></returns>
        public static T WithMaxRetry<T>(this T baseRequest, TimeSpan retriesTimeLimit) where T : IBaseRequest
        {
            string retryOptionKey = typeof(RetryHandlerOption).ToString();
            if (baseRequest.MiddlewareOptions.ContainsKey(retryOptionKey))
            {
                (baseRequest.MiddlewareOptions[retryOptionKey] as RetryHandlerOption).RetriesTimeLimit = retriesTimeLimit;
            }
            else
            {
                baseRequest.MiddlewareOptions.Add(retryOptionKey, new RetryHandlerOption { RetriesTimeLimit = retriesTimeLimit });
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets the maximum number of redirects to the default Redirect Middleware Handler for this request.
        /// This only works with the default Redirect Middleware Handler.
        /// If you use a custom Redirect Middleware Handler, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseRequest">The <see cref="BaseRequest"/> for the request.</param>
        /// <param name="maxRedirects">Maximum number of redirects allowed for the request</param>
        /// <returns></returns>
        public static T WithMaxRedirects<T>(this T baseRequest, int maxRedirects) where T : IBaseRequest
        {
            string redirectOptionKey = typeof(RedirectHandlerOption).ToString();
            if (baseRequest.MiddlewareOptions.ContainsKey(redirectOptionKey))
            {
                (baseRequest.MiddlewareOptions[redirectOptionKey] as RedirectHandlerOption).MaxRedirect = maxRedirects;
            }
            else
            {
                baseRequest.MiddlewareOptions.Add(redirectOptionKey, new RedirectHandlerOption { MaxRedirect = maxRedirects });
            }
            return baseRequest;
        }

        /// <summary>
        /// Sets Microsoft Graph's scopes that will be used by <see cref="IAuthenticationProvider"/> to authenticate this request
        /// and can be used to perform incremental scope consent.
        /// This only works with the default authentication handler and default set of Microsoft graph authentication providers.
        /// If you use a custom authentication handler or authentication provider, you have to handle it's retrieval in your implementation.
        /// </summary>
        /// <param name="baseRequest">The <see cref="IBaseRequest"/>.</param>
        /// <param name="scopes">Microsoft graph scopes used to authenticate this request.</param>
        public static T WithScopes<T>(this T baseRequest, string[] scopes) where T : IBaseRequest
        {
            string authHandlerOptionKey = typeof(AuthenticationHandlerOption).ToString();
            AuthenticationHandlerOption authHandlerOptions = baseRequest.MiddlewareOptions[authHandlerOptionKey] as AuthenticationHandlerOption;
            AuthenticationProviderOption msalAuthProviderOption = authHandlerOptions.AuthenticationProviderOption as AuthenticationProviderOption ?? new AuthenticationProviderOption();

            msalAuthProviderOption.Scopes = scopes;

            authHandlerOptions.AuthenticationProviderOption = msalAuthProviderOption;
            baseRequest.MiddlewareOptions[authHandlerOptionKey] = authHandlerOptions;

            return baseRequest;
        }

        /// <summary>
        /// Sets MSAL's force refresh flag to <see cref="IAuthenticationProvider"/> for this request. If set to true, <see cref="IAuthenticationProvider"/> will refresh existing access token in cahce.
        /// This defaults to false if not set.
        /// </summary>
        /// <param name="baseRequest">The <see cref="IBaseRequest"/>.</param>
        /// <param name="forceRefresh">A <see cref="bool"/> flag to determine whether refresh access token or not.</param>
        public static T WithForceRefresh<T>(this T baseRequest, bool forceRefresh) where T : IBaseRequest
        {
            string authHandlerOptionKey = typeof(AuthenticationHandlerOption).ToString();
            AuthenticationHandlerOption authHandlerOptions = baseRequest.MiddlewareOptions[authHandlerOptionKey] as AuthenticationHandlerOption;
            AuthenticationProviderOption msalAuthProviderOption = authHandlerOptions.AuthenticationProviderOption as AuthenticationProviderOption ?? new AuthenticationProviderOption();

            msalAuthProviderOption.ForceRefresh = forceRefresh;

            authHandlerOptions.AuthenticationProviderOption = msalAuthProviderOption;
            baseRequest.MiddlewareOptions[authHandlerOptionKey] = authHandlerOptions;

            return baseRequest;
        }

        /// <summary>
        /// Sets <see cref="GraphUserAccount"/> to be used to retrieve an access token for this request.
        /// It is also used to handle multi-user/ multi-tenant access token cache storage and retrieval.
        /// </summary>
        /// <param name="baseRequest">The <see cref="IBaseRequest"/>.</param>
        /// <param name="userAccount">A <see cref="GraphUserAccount"/> that represents a user account. At a minimum, the ObjectId and TenantId must be set.
        /// </param>
        public static T WithUserAccount<T>(this T baseRequest, GraphUserAccount userAccount) where T : IBaseRequest
        {
            string authHandlerOptionKey = typeof(AuthenticationHandlerOption).ToString();
            AuthenticationHandlerOption authHandlerOptions = baseRequest.MiddlewareOptions[authHandlerOptionKey] as AuthenticationHandlerOption;
            AuthenticationProviderOption msalAuthProviderOption = authHandlerOptions.AuthenticationProviderOption as AuthenticationProviderOption ?? new AuthenticationProviderOption();

            msalAuthProviderOption.UserAccount = userAccount;

            authHandlerOptions.AuthenticationProviderOption = msalAuthProviderOption;
            baseRequest.MiddlewareOptions[authHandlerOptionKey] = authHandlerOptions;

            return baseRequest;
        }

#if !iOS // Don't make this available for iOS mobile as it can't/shouldn't use confidential clients
        /// <summary>
        /// Sets <see cref="UserAssertion"/> for this request.
        /// This should only be used with <see cref="OnBehalfOfProvider"/>.
        /// </summary>
        /// <param name="baseRequest">The <see cref="IBaseRequest"/>.</param>
        /// <param name="userAssertion">A <see cref="UserAssertion"/> for the user.</param>
        public static T WithUserAssertion<T>(this T baseRequest, UserAssertion userAssertion) where T : IBaseRequest
        {
            string authHandlerOptionKey = typeof(AuthenticationHandlerOption).ToString();
            AuthenticationHandlerOption authHandlerOptions = baseRequest.MiddlewareOptions[authHandlerOptionKey] as AuthenticationHandlerOption;
            AuthenticationProviderOption msalAuthProviderOption = authHandlerOptions.AuthenticationProviderOption as AuthenticationProviderOption ?? new AuthenticationProviderOption();

            msalAuthProviderOption.UserAssertion = userAssertion;

            authHandlerOptions.AuthenticationProviderOption = msalAuthProviderOption;
            baseRequest.MiddlewareOptions[authHandlerOptionKey] = authHandlerOptions;

            return baseRequest;
        }
#endif

        /// <summary>
        /// Sets a username (email) and password of an Azure AD account to authenticate.
        /// This should only be used with <see cref="UsernamePasswordProvider"/>.
        /// This provider is NOT RECOMMENDED because it exposes the users password.
        /// We recommend you use <see cref="IntegratedWindowsAuthenticationProvider"/> instead.
        /// </summary>
        /// <param name="baseRequest">The <see cref="IBaseRequest"/>.</param>
        /// <param name="email">Email address of the user to authenticate.</param>
        /// <param name="password">Password of the user to authenticate.</param>
        public static T WithUsernamePassword<T>(this T baseRequest, string email, SecureString password) where T : IBaseRequest
        {
            string authHandlerOptionKey = typeof(AuthenticationHandlerOption).ToString();
            AuthenticationHandlerOption authHandlerOptions = baseRequest.MiddlewareOptions[authHandlerOptionKey] as AuthenticationHandlerOption;
            AuthenticationProviderOption msalAuthProviderOption = authHandlerOptions.AuthenticationProviderOption as AuthenticationProviderOption ?? new AuthenticationProviderOption();

            msalAuthProviderOption.Password = password;
            msalAuthProviderOption.UserAccount = new GraphUserAccount { Email = email };

            authHandlerOptions.AuthenticationProviderOption = msalAuthProviderOption;
            baseRequest.MiddlewareOptions[authHandlerOptionKey] = authHandlerOptions;

            return baseRequest;
        }
    }
}
