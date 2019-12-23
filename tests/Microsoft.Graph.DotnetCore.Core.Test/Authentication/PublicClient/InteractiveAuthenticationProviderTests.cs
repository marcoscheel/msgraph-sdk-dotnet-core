// ------------------------------------------------------------------------------
//  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

namespace Microsoft.Graph.DotnetCore.Core.Test.Authentication.PublicClient
{
#if !iOS && !ANDROID
    using Microsoft.Identity.Client;
    using Moq;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Xunit;

    public class InteractiveAuthenticationProviderTests
    {
        [Fact]
        public void ShouldConstructAuthProviderWithPublicClientApp()
        {
            string clientId = "00000000-0000-0000-0000-000000000000";
            string authority = "https://login.microsoftonline.com/organizations/";
            IEnumerable<string> scopes = new List<string> { "User.ReadBasic.All" };

            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(authority)
                .Build();

            InteractiveAuthenticationProvider auth = new InteractiveAuthenticationProvider(publicClientApplication, scopes);

            Assert.IsAssignableFrom<IAuthenticationProvider>(auth);
            Assert.NotNull(auth.ClientApplication);
            Assert.Equal(publicClientApplication, auth.ClientApplication);
        }

        [Fact]
        public void ConstructorShouldThrowExceptionWithNullPublicClientApp()
        {
            IEnumerable<string> scopes = new List<string> { "User.ReadBasic.All" };

            AuthenticationException ex = Assert.Throws<AuthenticationException>(() => new InteractiveAuthenticationProvider(null, scopes));

            Assert.Equal(ex.Error.Code, ErrorConstants.Codes.InvalidRequest);
            Assert.Equal(ex.Error.Message, String.Format(ErrorConstants.Messages.NullValue, "publicClientApplication"));
        }

        [Fact]
        public void ShouldUseDefaultScopeUrlWhenScopeIsNull()
        {
            var mock = Mock.Of<IPublicClientApplication>();

            InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(mock, null);

            Assert.NotNull(authProvider.Scopes);
            Assert.True(authProvider.Scopes.Count().Equals(1));
            Assert.Equal(AuthConstants.DefaultScopeUrl, authProvider.Scopes.FirstOrDefault());
        }

        [Fact]
        public void ShouldThrowExceptionWhenScopesAreEmpty()
        {
            var mock = Mock.Of<IPublicClientApplication>();

            AuthenticationException ex = Assert.Throws<AuthenticationException>(() => new InteractiveAuthenticationProvider(mock, Enumerable.Empty<string>()));

            Assert.Equal(ex.Error.Message, ErrorConstants.Messages.EmptyScopes);
            Assert.Equal(ex.Error.Code, ErrorConstants.Codes.InvalidRequest);
        }
    }
#endif
}