import { IInputs, IOutputs } from "./generated/ManifestTypes";

import 'jquery';
import 'popper.js';
import 'bootstrap';
import * as FileSaver from 'file-saver';

class EntityReference {
	id: string;
	typeName: string;
	constructor(typeName: string, id: string) {
		this.id = id;
		this.typeName = typeName;
	}
}

class AttachedFile implements ComponentFramework.FileObject {
	fileContent: string;
	fileSize: number;
	fileName: string;
	mimeType: string;
	constructor(fileName: string, mimeType: string, fileContent: string, fileSize: number) {
		this.fileName = fileName;
		this.mimeType = mimeType;
		this.fileContent = fileContent;
		this.fileSize = fileSize;
	}
}

export class CarouselControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _indicatorList: HTMLUListElement;
	private _slides: HTMLDivElement;
	private _carousel: HTMLDivElement;
	private _container: HTMLDivElement;

	private _context: ComponentFramework.Context<IInputs>;

	private _imageHeight: number | null;
	private _imageWidth: number | null;
	private entityReference: EntityReference;

	private _showIndicators: boolean;
	private _showArrows: boolean;
	private _showFilename: boolean;
	private _showSlideAnimation: boolean;

	private _supportedMimeTypes: string[] = ["image/jpeg", "image/png", "image/svg+xml"];
	private _supportedExtensions : string[] = [".jpg", ".jpeg", ".png", ".svg", ".gif"];

	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._context = context;
		this._container = container;

		this.entityReference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		)

		this._imageHeight = context.parameters.height != undefined ? context.parameters.height.raw : null;
		this._imageWidth = context.parameters.width != undefined ? context.parameters.width.raw : null;

		this._showArrows = this.GetShowHideValue(context.parameters.showArrows);
		this._showIndicators = this.GetShowHideValue(context.parameters.showIndicators);
		this._showFilename = this.GetShowHideValue(context.parameters.showFilename);
		this._showSlideAnimation = this.GetShowHideValue(context.parameters.showSlideAnimation);

		let carousel = document.createElement("div");
		carousel.id = "demo";
		carousel.classList.add("carousel");
		if (this._showSlideAnimation) { carousel.classList.add("slide"); }
		carousel.setAttribute("data-ride", "carousel");

		this._carousel = carousel;

		let indicatorList = document.createElement("ul");
		indicatorList.classList.add("carousel-indicators");

		this._indicatorList = indicatorList;

		let slides = document.createElement("div");
		slides.classList.add("carousel-inner");

		this._slides = slides;

		carousel.appendChild(indicatorList);
		carousel.appendChild(slides);
		if (this._showArrows) { this.AddNavigation(); }
		this.GetFiles(this.entityReference).then(result => this.BuildCarousel(result));
		
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {		
		
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}

	private AddIndicator(index: number): void {
		let liElement = document.createElement("li");
		if (index == 0) { liElement.classList.add("active"); }
		liElement.setAttribute("data-target", "#demo");
		liElement.setAttribute("data-slide-to", index.toString());
		this._indicatorList.appendChild(liElement);
	}

	private AddSlide(index: number, file: AttachedFile): void {
		let source = 'data:' + file.mimeType + ';base64, ' + file.fileContent;
		let carouselItem = document.createElement("div");
		carouselItem.classList.add("carousel-item");
		if (index == 0) { carouselItem.classList.add("active"); }

		let imgElement = document.createElement("img");
		imgElement.src = source;
		if (this._imageHeight != null) { imgElement.height = this._imageHeight }
		if (this._imageWidth != null) { imgElement.width = this._imageWidth }

		carouselItem.appendChild(imgElement);

		if(this._showFilename){
			let imageCaption = document.createElement("div");
			imageCaption.classList.add("carousel-caption");
	
			let fileName = document.createElement("h3");
			fileName.textContent = file.fileName;
			fileName.onclick = (e => { this.DownloadFile(file); });
	
			imageCaption.appendChild(fileName);
			carouselItem.appendChild(imageCaption);
		}

		this._slides.appendChild(carouselItem);
	}

	private AddNavigation(): void {
		let leftArrowElement = document.createElement("a");
		leftArrowElement.classList.add("carousel-control-prev");
		leftArrowElement.href = "#demo";
		leftArrowElement.setAttribute("data-slide", "prev");

		let leftSpan = document.createElement("span");
		leftSpan.classList.add("carousel-control-prev-icon");

		leftArrowElement.appendChild(leftSpan);

		let rightArrowElement = document.createElement("a");
		rightArrowElement.classList.add("carousel-control-next");
		rightArrowElement.href = "#demo";
		rightArrowElement.setAttribute("data-slide", "next");

		let rightSpan = document.createElement("span");
		rightSpan.classList.add("carousel-control-next-icon");

		rightArrowElement.appendChild(rightSpan);

		this._carousel.appendChild(leftArrowElement);
		this._carousel.appendChild(rightArrowElement);
	}

	private async GetFiles(ref: EntityReference): Promise<AttachedFile[]> {
		let attachmentType = ref.typeName == "email" ? "activitymimeattachment" : "annotation";
		let fetchXml =
			"<fetch>" +
				"<entity name='" + attachmentType + "'>" +
					"<filter>" +
						"<condition attribute='objectid' operator='eq' value='" + ref.id + "'/>" +
						"<condition attribute='filesize' operator='gt' value='0'/>" +
					"</filter>" +
				"</entity>" +
			"</fetch>";

		let query = '?fetchXml=' + encodeURIComponent(fetchXml);

		try {
			const result = await this._context.webAPI.retrieveMultipleRecords(attachmentType, query);
			let items = [];
			for (let i = 0; i < result.entities.length; i++) {
				let record = result.entities[i];
				let fileName = <string>record["filename"];
				let mimeType = <string>record["mimetype"];
				let content = <string>record["body"] || <string>record["documentbody"];
				let fileSize = <number>record["filesize"];

				const ext = fileName.substr(fileName.lastIndexOf('.')).toLowerCase();

				if (!this._supportedMimeTypes.includes(mimeType) && !this._supportedExtensions.includes(ext)) { continue; }

				let file = new AttachedFile(fileName, mimeType, content, fileSize);
				items.push(file);
			}
			return items;
		}
		catch (error) {
			console.log(error);
			return [];
		}
	}

	private Base64ToFile(base64Data: string, tempfilename: string, contentType: string): File {
		contentType = contentType || '';
		const sliceSize = 1024;
		const byteCharacters = atob(base64Data);
		const bytesLength = byteCharacters.length;
		const slicesCount = Math.ceil(bytesLength / sliceSize);
		const byteArrays = new Array(slicesCount);

		for (let sliceIndex = 0; sliceIndex < slicesCount; ++sliceIndex) {
			const begin = sliceIndex * sliceSize;
			const end = Math.min(begin + sliceSize, bytesLength);

			const bytes = new Array(end - begin);
			for (let offset = begin, i = 0; offset < end; ++i, ++offset) {
				bytes[i] = byteCharacters[offset].charCodeAt(0);
			}
			byteArrays[sliceIndex] = new Uint8Array(bytes);
		}
		return new File(byteArrays, tempfilename, { type: contentType });
	}

	private DownloadFile(file: AttachedFile): void {
		const myFile = this.Base64ToFile(file.fileContent, file.fileName, file.mimeType);
		FileSaver.saveAs(myFile, file.fileName);
	}

	private BuildCarousel(files: AttachedFile[]): void {
		for (let index = 0; index < files.length; index++) {
			const file = files[index];

			if (this._showIndicators) { this.AddIndicator(index);}
			this.AddSlide(index, file);
		}

		if (files.length > 0) {
			this._container.appendChild(this._carousel);
		}
	}

	private GetShowHideValue(value: ComponentFramework.PropertyTypes.StringProperty | undefined): boolean {
		if (value == undefined) {
			return true;
		} else {
			if (value.raw == "1" || value.raw == "" || value.raw == null) {
				return true;
			} else if (value.raw == "0") {
				return false;
			}
			return false;
		}
	}
}