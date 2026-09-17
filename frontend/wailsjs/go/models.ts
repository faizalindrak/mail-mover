export namespace main {
	
	export class EnvironmentStatus {
	    windows: boolean;
	    outlookRunning: boolean;
	    pstNames: string[];
	    message: string;
	
	    static createFrom(source: any = {}) {
	        return new EnvironmentStatus(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.windows = source["windows"];
	        this.outlookRunning = source["outlookRunning"];
	        this.pstNames = source["pstNames"];
	        this.message = source["message"];
	    }
	}
	export class MigrationRequest {
	    sourcePath: string;
	    stagingPath: string;
	    targetPSTName: string;
	    markAsRead: boolean;
	    preventSleep: boolean;
	    skipConversion: boolean;
	    keepStagingEML: boolean;
	
	    static createFrom(source: any = {}) {
	        return new MigrationRequest(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.sourcePath = source["sourcePath"];
	        this.stagingPath = source["stagingPath"];
	        this.targetPSTName = source["targetPSTName"];
	        this.markAsRead = source["markAsRead"];
	        this.preventSleep = source["preventSleep"];
	        this.skipConversion = source["skipConversion"];
	        this.keepStagingEML = source["keepStagingEML"];
	    }
	}

}

