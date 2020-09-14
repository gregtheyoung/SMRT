﻿using System;
using System.Diagnostics.CodeAnalysis;
using FakeItEasy;
using Testinvi.Helpers;
using Tweetinvi.Core.Credentials;
using Tweetinvi.Models;

namespace Testinvi.SetupHelpers
{
    [ExcludeFromCodeCoverage]
    public static class CredentialsAccessorHelper
    {
        public static void SetupPassThrough(this Fake<ICredentialsAccessor> fakeCredentialsAccessor)
        {
            fakeCredentialsAccessor
                .CallsTo(x => x.ExecuteOperationWithCredentials(It.IsAny<ITwitterCredentials>(), It.IsAny<Action>()))
                .Invokes(x =>
                {
                    x.Arguments.Get<Action>(1).Invoke();
                });
        }

        public static void SetupPassThrough<T>(this Fake<ICredentialsAccessor> fakeCredentialsAccessor) where T : class 
        {
            fakeCredentialsAccessor
                .CallsTo(x => x.ExecuteOperationWithCredentials(It.IsAny<ITwitterCredentials>(), It.IsAny<Func<T>>()))
                .ReturnsLazily((ITwitterCredentials cred, Func<T> f) => f());
        }
    }
}
