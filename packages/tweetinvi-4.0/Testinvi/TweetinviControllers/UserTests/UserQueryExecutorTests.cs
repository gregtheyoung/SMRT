﻿using System.Collections.Generic;
using System.IO;
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Testinvi.Helpers;
using Testinvi.SetupHelpers;
using Tweetinvi.Controllers.User;
using Tweetinvi.Core.Helpers;
using Tweetinvi.Core.Parameters;
using Tweetinvi.Core.QueryGenerators;
using Tweetinvi.Core.Web;
using Tweetinvi.Models;
using Tweetinvi.Models.DTO;
using Tweetinvi.Models.DTO.QueryDTO;

namespace Testinvi.TweetinviControllers.UserTests
{
    [TestClass]
    public class UserQueryExecutorTests
    {
        private FakeClassBuilder<UserQueryExecutor> _fakeBuilder;
        private Fake<IUserQueryGenerator> _fakeUserQueryGenerator;
        private Fake<ITwitterAccessor> _fakeTwitterAccessor;
        private Fake<IWebHelper> _fakeWebHelper;

        private List<long> _cursorQueryIds;

        [TestInitialize]
        public void TestInitialize()
        {
            _fakeBuilder = new FakeClassBuilder<UserQueryExecutor>();
            _fakeUserQueryGenerator = _fakeBuilder.GetFake<IUserQueryGenerator>();
            _fakeTwitterAccessor = _fakeBuilder.GetFake<ITwitterAccessor>();
            _fakeWebHelper = _fakeBuilder.GetFake<IWebHelper>();

            _cursorQueryIds = new List<long>();
        }

        #region FriendIds

        // This tests that if the CursorQuery returns null, the accessor returns null
        [TestMethod]
        public void GetFriendIdWithUserDTOs_TwitterAccessorReturnsNull_ReturnsNull()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var maximumNumberOfFriends = TestHelper.GenerateRandomInt();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetFriendIdsQuery(userDTO, maximumNumberOfFriends)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeExecuteCursorGETQuery<long, IIdsCursorQueryResultDTO>(expectedQuery, null);

            // Act
            var result = queryExecutor.GetFriendIds(userDTO, maximumNumberOfFriends);

            // Assert
            Assert.IsNull(result);
        }
        [TestMethod]
        public void GetFriendIdsWithUserDTO_AnyData_ReturnsTwitterAccessorResult()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var maximumNumberOfFriends = TestHelper.GenerateRandomInt();
            var expectedQuery = TestHelper.GenerateString();
            var expectedCursorResults = GenerateExpectedCursorResults();

            _fakeUserQueryGenerator.CallsTo(x => x.GetFriendIdsQuery(userDTO, maximumNumberOfFriends)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeExecuteCursorGETQuery<long, IIdsCursorQueryResultDTO>(expectedQuery, expectedCursorResults);

            // Act
            var result = queryExecutor.GetFriendIds(userDTO, maximumNumberOfFriends);

            // Assert
            Assert.IsTrue(result.ContainsAll(_cursorQueryIds));
        }

        #endregion

        #region FollowerIds

        // This tests that if the CursorQuery returns null, the accessor returns null
        [TestMethod]
        public void GetFollowerIdWithUserDTOs_TwitterAccessorReturnsNull_ReturnsNull()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var maximumNumberOfFollowers = TestHelper.GenerateRandomInt();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetFollowerIdsQuery(userDTO, maximumNumberOfFollowers)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeExecuteCursorGETQuery<long, IIdsCursorQueryResultDTO>(expectedQuery, null);

            // Act
            var result = queryExecutor.GetFollowerIds(userDTO, maximumNumberOfFollowers);

            // Assert
            Assert.IsNull(result);
        }
        [TestMethod]
        public void GetFollowerIdsWithUserDTO_AnyData_ReturnsTwitterAccessorResult()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var maximumNumberOfFollowers = TestHelper.GenerateRandomInt();
            var expectedQuery = TestHelper.GenerateString();
            var expectedCursorResults = GenerateExpectedCursorResults();

            _fakeUserQueryGenerator.CallsTo(x => x.GetFollowerIdsQuery(userDTO, maximumNumberOfFollowers)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeExecuteCursorGETQuery<long, IIdsCursorQueryResultDTO>(expectedQuery, expectedCursorResults);

            // Act
            var result = queryExecutor.GetFollowerIds(userDTO, maximumNumberOfFollowers);

            // Assert
            Assert.IsTrue(result.ContainsAll(_cursorQueryIds));
        }

        #endregion

        #region FavouriteTweets

        [TestMethod]
        public void GetFavouriteTweetsWithUserDTO_AnyData_ReturnsTwitterAccessorResult()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            TestHelper.GenerateRandomInt();
            var expectedQuery = TestHelper.GenerateString();
            IEnumerable<ITweetDTO> expectedResult = new[] { A.Fake<ITweetDTO>() };
            var parameters = It.IsAny<IGetUserFavoritesQueryParameters>();

            _fakeUserQueryGenerator.CallsTo(x => x.GetFavoriteTweetsQuery(parameters)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeExecuteGETQuery(expectedQuery, expectedResult);

            // Act
            var result = queryExecutor.GetFavoriteTweets(parameters);

            // Assert
            Assert.AreEqual(result, expectedResult);
        }

        #endregion

        #region Block User

        [TestMethod]
        public void BlockUser_WithUserDTO_ReturnsTrue()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetBlockUserQuery(userDTO)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeTryExecutePOSTQuery(expectedQuery, true);

            // Act
            var result = queryExecutor.BlockUser(userDTO);

            // Assert
            Assert.AreEqual(result, true);
        }

        [TestMethod]
        public void BlockUser_WithUserDTO_ReturnsFalse()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetBlockUserQuery(userDTO)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeTryExecutePOSTQuery(expectedQuery, false);

            // Act
            var result = queryExecutor.BlockUser(userDTO);

            // Assert
            Assert.AreEqual(result, false);
        }

        #endregion

        #region Stream Profile Image

        [TestMethod]
        public void GetProfileImageStream_ReturnsWebHelperResult()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var stream = A.Fake<Stream>();
            var userDTO = A.Fake<IUserDTO>();
            var url = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.DownloadProfileImageURL(userDTO, ImageSize.bigger)).Returns(url);
            _fakeWebHelper.CallsTo(x => x.GetResponseStream(url)).Returns(stream);

            // Act
            var result = queryExecutor.GetProfileImageStream(userDTO, ImageSize.bigger);

            // Assert
            Assert.AreEqual(result, stream);
        }

        #endregion

        #region Spam

        [TestMethod]
        public void ReportUserForSpam_WithUserDTO_ReturnsTrue()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetReportUserForSpamQuery(userDTO)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeTryExecutePOSTQuery(expectedQuery, true);

            // Act
            var result = queryExecutor.ReportUserForSpam(userDTO);

            // Assert
            Assert.AreEqual(result, true);
        }

        [TestMethod]
        public void ReportUserForSpam_WithUserDTO_ReturnsFalse()
        {
            // Arrange
            var queryExecutor = CreateUserQueryExecutor();
            var userDTO = A.Fake<IUserDTO>();
            var expectedQuery = TestHelper.GenerateString();

            _fakeUserQueryGenerator.CallsTo(x => x.GetReportUserForSpamQuery(userDTO)).Returns(expectedQuery);
            _fakeTwitterAccessor.ArrangeTryExecutePOSTQuery(expectedQuery, false);

            // Act
            var result = queryExecutor.ReportUserForSpam(userDTO);

            // Assert
            Assert.AreEqual(result, false);
        }

        #endregion

        private IEnumerable<long> GenerateExpectedCursorResults()
        {
            var queryId1 = TestHelper.GenerateRandomLong();
            var queryId2 = TestHelper.GenerateRandomLong();

            _cursorQueryIds.Add(queryId1);
            _cursorQueryIds.Add(queryId2);

            return new[] {queryId1, queryId2};
        }

        public UserQueryExecutor CreateUserQueryExecutor()
        {
            return _fakeBuilder.GenerateClass();
        }
    }
}