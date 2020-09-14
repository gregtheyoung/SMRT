﻿using System.Collections.Generic;
using Tweetinvi.Models;
using Tweetinvi.Models.DTO;
using Tweetinvi.Parameters;

namespace Tweetinvi.Core.Controllers
{
    public interface ITweetController
    {
        // Publish Tweet
        ITweet PublishTweet(IPublishTweetParameters parameters);
        ITweet PublishTweet(string text, IPublishTweetOptionalParameters optionalParameters = null);

        ITweet PublishTweetInReplyTo(string text, long tweetId);
        ITweet PublishTweetInReplyTo(string text, ITweetIdentifier tweet);

        bool CanBePublished(IPublishTweetParameters publishTweetParameters);
        bool CanBePublished(string text, IPublishTweetOptionalParameters publishTweetOptionalParameters = null);

        // Publish Retweet
        ITweet PublishRetweet(ITweet tweetToPublish);
        ITweet PublishRetweet(ITweetDTO tweetToPublish);
        ITweet PublishRetweet(long tweetId);
        
        // Publish UnRetweet
        ITweet UnRetweet(ITweetIdentifier tweetToPublish);
        ITweet UnRetweet(long tweetId);

        // Get Retweets
        IEnumerable<ITweet> GetRetweets(ITweetIdentifier tweet, int maxRetweetsToRetrieve = 100);
        IEnumerable<ITweet> GetRetweets(long tweetId, int maxRetweetsToRetrieve = 100);

        // Get Retweeters
        IEnumerable<long> GetRetweetersIds(ITweetIdentifier tweet, int maxRetweetersToRetrieve = 100);
        IEnumerable<long> GetRetweetersIds(long tweetId, int maxRetweetersToRetrieve = 100);

        // Destroy Tweet
        bool DestroyTweet(ITweet tweet);
        bool DestroyTweet(ITweetDTO tweetDTO);
        bool DestroyTweet(long tweetId);

        // Generate OembedTweet
        IOEmbedTweet GenerateOEmbedTweet(ITweet tweet);
        IOEmbedTweet GenerateOEmbedTweet(ITweetDTO tweetDTO);
        IOEmbedTweet GenerateOEmbedTweet(long tweetId);

        // Favorite Tweet
        bool FavoriteTweet(ITweet tweet);
        bool FavoriteTweet(ITweetDTO tweetDTO);
        bool FavoriteTweet(long tweetId);

        bool UnFavoriteTweet(ITweet tweet);
        bool UnFavoriteTweet(ITweetDTO tweetDTO);
        bool UnFavoriteTweet(long tweetId);

        // Update Published Tweet
        void UpdateTweetIfTweetSuccessfullyBeenPublished(ITweet sourceTweet, ITweetDTO publishedTweetDTO);
        void UploadMedias(IPublishTweetParameters parameters);
    }
}