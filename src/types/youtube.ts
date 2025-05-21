export interface ChannelInfo {
	title: string;
	subscriberCount: string;
}

export interface VideoItem {
	id: {
		videoId: string;
	};
	snippet: {
		title: string;
		description: string;
		channelId: string;
		channelTitle: string;
		publishedAt: string;
	};
	statistics?: {
		viewCount: string;
		likeCount: string;
		dislikeCount: string;
		commentCount: string;
	};
}

export interface SearchResponse {
	items: Array<{
		id: {
			videoId: string;
		};
		snippet: {
			title: string;
			description: string;
			channelId: string;
			channelTitle: string;
			publishedAt: string;
		};
	}>;
}

export interface VideosResponse {
	items: VideoItem[];
}

export interface ChannelsResponse {
	items: Array<{
		id: string;
		snippet: {
			title: string;
		};
		statistics: {
			subscriberCount: string;
		};
	}>;
}
