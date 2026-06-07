import { Router, type IRouter } from "express";
import healthRouter from "./health";
import warehouseRouter from "./warehouse";

const router: IRouter = Router();

router.use(healthRouter);
router.use("/warehouse", warehouseRouter);

export default router;
