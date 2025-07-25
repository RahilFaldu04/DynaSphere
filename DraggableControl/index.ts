import { IInputs, IOutputs } from "./generated/ManifestTypes";
import Sortable from 'sortablejs';


interface Project {
    dyn_projectid: string;
    dyn_name: string;
}

interface ProjectTask {
    dyn_projecttaskid: string;
    dyn_taskname: string;
    dyn_status: number;
    assignto: string;
    dyn_projecttask: string;
}

export class DraggableControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;
    private dropdownData: Project[] = [];
    private tasks: ProjectTask[] = [];
    private notifyOutputChanged: () => void;

    private readonly STATUS_MAP: Record<number, string> = {
        1: "Not Started",
        2: "In Progress",
        4: "In Review",
        3: "Completed"
    };

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        const cleanUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
        this.loadProjects(cleanUserId).then(() => this.renderDropdown()).catch(console.error);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        // Cleanup if needed
    }


    private async loadProjects(cleanUserId: string): Promise<void> {
        // Load all Projects and fill dropdown
        await this.context.webAPI.retrieveMultipleRecords(
            "dyn_resource",
            `?$select=dyn_resourceid&$filter=_dyn_user_value eq ${cleanUserId}`
        ).then(async resourceResult => {
            if (resourceResult.entities.length === 0) {
                console.log("No resource found for current user.");
                return;
            }
            const resourceIds = resourceResult.entities.map(r => r.dyn_resourceid);
            const projectIds: string[] = [];

            for (const resourceId of resourceIds) {
                const teamResult = await this.context.webAPI.retrieveMultipleRecords(
                    "dyn_projectteammember",
                    `?$select=_dyn_project_value&$filter=_dyn_resource_value eq ${resourceId}`
                );
                teamResult.entities.forEach(e => {
                    if (e._dyn_project_value) {
                        projectIds.push(e._dyn_project_value);
                    }
                });
            }

            if (projectIds.length === 0) {
                console.log("No projects found for this user's team memberships.");
                return;
            }

            const orConditions = projectIds.map(id => `(dyn_projectid eq ${id})`).join(" or ");
            const projectFilter = `?$filter=${orConditions}&$select=dyn_name,dyn_projectid`;
            const res = await this.context.webAPI.retrieveMultipleRecords("dyn_project", projectFilter);
            this.dropdownData = res.entities as Project[];
            return;
        })
    };

    private async loadTasks(projectId: string): Promise<void> {
        const res = await this.context.webAPI.retrieveMultipleRecords(
            "dyn_projecttask",
            `?$select=dyn_projecttaskid,dyn_taskname,dyn_status,_createdby_value,dyn_projecttask,modifiedon&$filter=_dyn_project_value eq ${projectId}&$orderby=modifiedon desc`
        );
        this.tasks = res.entities.map(task => ({
            dyn_projecttaskid: task.dyn_projecttaskid,
            dyn_taskname: task.dyn_taskname,
            dyn_status: task.dyn_status,
            assignto: task["_createdby_value@OData.Community.Display.V1.FormattedValue"],
            dyn_projecttask: task.dyn_projecttask,
        }));
    }

    private renderDropdown(): void {
        this.container.innerHTML = "";
        this.container.style.fontFamily = `"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif`;

        const wrapper = document.createElement("div");
        wrapper.style.marginBottom = "20px";
        wrapper.style.display = "flex";
        wrapper.style.alignItems = "center";
        wrapper.style.gap = "10px";

        const label = document.createElement("label");
        label.innerText = "Project :";
        label.style.fontWeight = "bold";
        label.style.fontSize = "16px";

        const dropdown = document.createElement("select");
        dropdown.style.padding = "10px";
        dropdown.style.borderRadius = "8px";
        dropdown.style.border = "1px solid #ccc";
        dropdown.style.fontSize = "16px";
        dropdown.style.minWidth = "300px";

        dropdown.innerHTML = this.dropdownData.map((p, index) =>
            `<option value="${p.dyn_projectid}" ${index === 0 ? "selected" : ""}>${p.dyn_name}</option>`
        ).join("");

        dropdown.onchange = async () => {
            if (!dropdown.value) return;
            await this.loadTasks(dropdown.value);
            this.renderKanban();
        };

        wrapper.appendChild(label);
        wrapper.appendChild(dropdown);
        this.container.appendChild(wrapper);

        // Load and render the first project by default
        if (this.dropdownData.length > 0) {
            this.loadTasks(this.dropdownData[0].dyn_projectid).then(() => this.renderKanban()).catch((error) => {
                console.error("Error loading Projects options:", error);
            });;
        }
    }

    private renderKanban(): void {
        const existing = this.container.querySelector(".kanban-board");
        if (existing) existing.remove();

        const board = document.createElement("div");
        board.className = "kanban-board";
        board.style.display = "flex";
        board.style.gap = "20px";
        board.style.overflowX = "auto";

        const statusColors: Record<number, string> = {
            1: "#f5a623",
            2: "#4fc3f7",
            4: "#8F00FF",
            3: "#8bc34a"
        };

        const orderedStatuses = [1, 2, 4, 3];

        for (const statusNum of orderedStatuses) {
            const column = document.createElement("div");
            column.dataset.status = statusNum.toString();
            column.style.flex = "1";
            column.style.minWidth = "250px";
            column.style.border = "1px solid #ddd";
            column.style.borderRadius = "8px";
            column.style.backgroundColor = "#f8f8f8";
            column.style.display = "flex";
            column.style.flexDirection = "column";
            column.style.height = "500px";

            const header = document.createElement("div");
            header.innerText = this.STATUS_MAP[statusNum];
            header.style.backgroundColor = statusColors[statusNum];
            header.style.color = "white";
            header.style.padding = "10px";
            header.style.textAlign = "center";
            header.style.fontWeight = "bold";
            column.appendChild(header);

            const taskList = document.createElement("div");
            taskList.className = "task-list";
            taskList.dataset.status = statusNum.toString();
            taskList.style.flex = "1";
            taskList.style.overflowY = "auto";
            taskList.style.padding = "10px";

            const filteredTasks = this.tasks.filter(t => t.dyn_status === statusNum);
            for (const task of filteredTasks) {
                const card = document.createElement("div");
                card.className = "task-card";
                card.dataset.id = task.dyn_projecttaskid;
                card.dataset.status = statusNum.toString();

                card.style.border = "1px solid #ccc";
                card.style.borderLeft = `5px solid ${statusColors[statusNum]}`;
                card.style.borderRadius = "5px";
                card.style.padding = "8px";
                card.style.margin = "8px 0";
                card.style.backgroundColor = "white";
                card.style.transition = "all 0.2s ease-in-out";

                const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
                const recordUrl = `${orgUrl}/main.aspx?etn=dyn_projecttask&pagetype=entityrecord&id=${task.dyn_projecttaskid}`;

                card.innerHTML = `
                <strong>
                    <a href="${recordUrl}" target="_blank" style="text-decoration:none; color:inherit;">
                        ${task.dyn_projecttask ? task.dyn_projecttask + " - " : ""}${task.dyn_taskname}
                        <br/><small style="font-weight:normal;">${task.assignto}</small>
                    </a>
                </strong>
            `;

                taskList.appendChild(card);
            }
                // Add Task Button

                const addTaskBtn = document.createElement("button");
                addTaskBtn.innerText = "+ Add Task";
                addTaskBtn.style.background = "transparent";
                addTaskBtn.style.border = "none";
                addTaskBtn.style.color = "#28a745"; // Green
                addTaskBtn.style.cursor = "pointer";
                addTaskBtn.style.fontWeight = "bold";
                addTaskBtn.style.margin = "10px auto";
                addTaskBtn.style.display = "block";

                addTaskBtn.onclick = () => {
                    const projectId = (this.container.querySelector("select") as HTMLSelectElement)?.value;
                    if (!projectId) {
                        alert("Please select a project first.");
                        return;
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const entityOptions: any = {
                        entityName: "dyn_projecttask",
                        useQuickCreateForm: true
                    };
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const formParameters: any = {
                        "dyn_project@odata.bind": `/dyn_projects(${projectId})`,
                        "dyn_status": statusNum
                    };

                    Xrm.Navigation.openForm(entityOptions, formParameters).then(
                        (lookup) => {

                            this.loadTasks(projectId).then(() => {
                                this.renderKanban()
                            }
                        ).catch((error) => {
                                console.error("Error loading tasks:", error);
                            });
                            return;
                        },
                    ).catch((error) => {
                        console.error("Quick Create failed or closed:", error);
                    });
                };

                column.appendChild(addTaskBtn);
         

            column.appendChild(taskList);
            board.appendChild(column);

            // 🔄 Initialize Sortable
            Sortable.create(taskList, {
                group: "kanban-tasks",
                animation: 150,
                ghostClass: "sortable-ghost",
                onEnd: async (evt) => {
                    const taskId = evt.item?.dataset?.id;
                    const fromStatus = parseInt(evt.from.dataset.status || "0");
                    const toStatus = parseInt(evt.to.dataset.status || "0");

                    if (!taskId || isNaN(toStatus)) return;

                    const task = this.tasks.find(t => t.dyn_projecttaskid === taskId);

                    if (task && task.dyn_status !== toStatus) {
                        task.dyn_status = toStatus;
                        this.renderKanban();

                        try {
                            await this.updateTaskStatus(taskId, toStatus);
                        } catch (error) {
                            console.error("Failed to update task status:", error);
                            alert("Failed to update status. Please try again.");
                            task.dyn_status = fromStatus; // revert on failure
                            this.renderKanban();
                        }
                    }
                }
            });
        }

        this.container.appendChild(board)
        console.log("hsdfcfvgbhnfcvxcvii");
       
    }

    private async updateTaskStatus(taskId: string, newStatus: number): Promise<void> {
        try {
            await this.context.webAPI.updateRecord("dyn_projecttask", taskId, { dyn_status: newStatus });
        } catch (error) {
            console.error("Failed to update status:", error);
        }
    }
}
