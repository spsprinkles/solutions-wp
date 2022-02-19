import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SolutionsWebPartWebPart.module.scss';
import * as strings from 'SolutionsWebPartWebPartStrings';

export interface ISolutionsWebPartWebPartProps {
  solution: string;
}

// Solution Imports

import { Components, ContextInfo, Web } from "gd-sprest-bs";

export default class SolutionsWebPartWebPart extends BaseClientSideWebPart<ISolutionsWebPartWebPartProps> {

  public render(): void {
    // Initialize the library
    ContextInfo.setPageContext(this.context);

    // Display an alert
    Components.Alert({
      el: this.domElement,
      header: "Loading",
      content: "Loading the configuration",
      type: Components.AlertTypes.Primary
    });

    // Load the configuration
    this.loadConfig().then(
      // Success
      () => {
        // Clear the webpart
        while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

        // Validate the selected solution
        if (this.properties.solution) {
          // Get the solution
          let solution = this.getSolution(this.properties.solution).then(
            // Success
            solution => {
              if (solution) {
                // Ensure the render method exists
                if (solution.render) {
                  // Render the solution
                  solution.render(this.domElement, this.context);
                } else {
                  // Solution was found but no render method exists
                  Components.Alert({
                    el: this.domElement,
                    header: "Error",
                    content: "The selected solution was found but was unable to be loaded. Contact your administrator.",
                    type: Components.AlertTypes.Danger
                  });
                }
              } else {
                // Solution wasn't found in the configuration
                Components.Alert({
                  el: this.domElement,
                  header: "Error",
                  content: "The selected solution was not found in the configuration.",
                  type: Components.AlertTypes.Danger
                });
              }
            },

            // Error
            msg => {
              // Solution wasn't found in the configuration
              Components.Alert({
                el: this.domElement,
                header: "Error",
                content: msg,
                type: Components.AlertTypes.Danger
              });
            }
          );
        } else {
          // Render a button
          let elButton = Components.Button({
            text: "Select Solution",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
              // Open the property pane
              this.context.propertyPane.open();
            }
          });

          // Render the content element
          let elContent = document.createElement("div");
          elContent.innerHTML = "<span>No solution has been selected.</span><br/>";
          elContent.appendChild(elButton.el);

          // No solution selected
          Components.Alert({
            el: this.domElement,
            header: "Success",
            content: elContent,
            type: Components.AlertTypes.Success
          });
        }
      },
      // Error
      msg => {
        // Clear the webpart
        while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

        // Render an alert
        Components.Alert({
          el: this.domElement,
          header: "Error",
          content: msg,
          type: Components.AlertTypes.Danger
        });
      }
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onPropertyPaneConfigurationStart(): void {
    // See if the property hasn't been initialized
    if (!this._propertyInit && this._solutions != null) {
      // Refresh the property pane
      this.context.propertyPane.refresh();

      // Set the flag
      this._propertyInit = true;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('solution', {
                  label: strings.SolutionFieldLabel,
                  options: this._solutions,
                  disabled: this._solutions == null
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Configuration File
   */

  private _config: {
    sourceUrl: string;
    solutions: {
      global: string;
      name: string;
      url: string;
    }[]
  } = null;
  private _propertyInit = false;
  private _solutions: {
    key: string;
    text: string;
  }[] = null;

  // URL to the configuration file
  private _webUrl = "/";
  get SourceUrl(): string { return this._webUrl + "clientsideassets/solutions/"; }
  get ConfigUrl(): string { return this.SourceUrl + "config.json"; }

  // Gets a solution by name
  private getSolution(key: string): PromiseLike<any> {
    // Return a promise
    return new Promise((resolve, reject) => {
      // Parse the solutions
      for (let i = 0; i < this._config.solutions.length; i++) {
        let solution = this._config.solutions[i];

        // See if this is the target solution
        if (solution.global == key) {
          // See if the solution exists
          let globalVar = window[key];
          if (globalVar) {
            // Resolve the request
            resolve(globalVar);
            return;
          }

          // Load the solution
          let elScript = document.createElement("script");
          elScript.id = solution.global;
          elScript.src = this._config.sourceUrl + "/" + solution.url;
          this.domElement.appendChild(elScript);

          // Wait for the solution to load
          let maxRetries = 10;
          let loopId = setInterval(() => {
            // See if the script is loaded
            if (globalVar = window[key]) {
              // Resolve the request
              resolve(globalVar);

              // Stop the loop
              clearInterval(loopId);
            }
            // See if we have maxed out the retries
            else if (--maxRetries == 0) {
              // Stop the interval
              clearInterval(loopId);

              // Reject the request
              reject("The script for this solution could not be loaded, or exceeded the wait time.");
            }
          }, 500);
        }
      }
    });
  }

  // Loads the configuration file
  private loadConfig(): PromiseLike<void> {
    return new Promise((resolve, reject) => {
      // Load the configuration file
      Web(this._webUrl).getFileByServerRelativeUrl(this.ConfigUrl).content().execute(
        // Success
        content => {
          // Convert the string to a json object
          try {
            // Set the configuration file
            this._config = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(content)));

            // Validate the source url property
            if ((this._config.sourceUrl || "").trim().length == 0) {
              // Invalid
              reject("The 'sourceUrl' property is not defined in the configuration.");
              return;
            }

            // Validate the solutions property
            if (this._config.solutions == null || typeof (this._config.solutions.length) !== "number") {
              // Invalid
              reject("The 'solutions' property could not be found or is invalid.");
              return;
            }

            // Parse the solutions
            let solutions = [];
            for (let i = 0; i < this._config.solutions.length; i++) {
              let solution = this._config.solutions[i];

              // Add the solution
              solutions.push({
                key: solution.global,
                text: solution.name
              });
            }

            // Set the solutions
            this._solutions = solutions;

            // See if the webpart is being edited
            if (this.context.propertyPane.isPropertyPaneOpen()) {
              // Refresh the property pane
              this.context.propertyPane.refresh();

              // Set the flag
              this._propertyInit = true;
            }

            // Resolve the request
            resolve();
          }
          catch {
            // Reject the request
            reject("The configuration file at '" + this.ConfigUrl + "' is invalid.");
          }
        },

        // Error
        () => {
          // Reject the request
          reject("The configuration file could not be found at '" + this.ConfigUrl + "'.");
        }
      );
    });
  }
}
