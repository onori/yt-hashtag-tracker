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
		tags?: string[];
	};
	statistics?: {
		viewCount: string;
		likeCount: string;
		dislikeCount: string;
		commentCount: string;
	};
	analytics?: {
		averageViewPercentage?: number; // 平均再生率 (%)
		averageViewDuration?: string; // 平均視聴時間 (ISO 8601 duration)
		impressions?: number; // インプレッション数
		impressionsClickThroughRate?: number; // インプレッションクリック率 (%)
		subscribersGained?: number; // 新規チャンネル登録者数
		retentionRate?: number[]; // 視聴維持率の配列 (0-100%)
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
