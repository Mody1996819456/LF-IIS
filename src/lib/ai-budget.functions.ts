import { createServerFn } from "@tanstack/react-start";
import { generateText } from "ai";
import { z } from "zod";
import { createLovableAiGatewayProvider } from "./ai-gateway.server";

const InputSchema = z.object({
  sheetLabel: z.string(),
  currentTotalCost: z.number(),
  currentTotalQty: z.number(),
  itemsCount: z.number(),
  topItems: z.array(
    z.object({
      name: z.string(),
      cost: z.number(),
      qty: z.number(),
    }),
  ),
  monthlyTotals: z.array(z.object({ name: z.string(), total: z.number() })),
});

export const analyzeBudgetForecast = createServerFn({ method: "POST" })
  .inputValidator((input: unknown) => InputSchema.parse(input))
  .handler(async ({ data }) => {
    const key = process.env.LOVABLE_API_KEY;
    if (!key) throw new Error("Missing LOVABLE_API_KEY");

    const gateway = createLovableAiGatewayProvider(key);

    const forecast20 = data.currentTotalCost * 1.2;
    const increase = forecast20 - data.currentTotalCost;

    const prompt = `أنت خبير مالي متخصص في تحليل موازنات المزارع والمنشآت الزراعية.
حلل بيانات الموازنة التالية للقسم "${data.sheetLabel}" وقدم تقريراً احترافياً باللغة العربية.

📊 البيانات الحالية:
- عدد الأصناف: ${data.itemsCount}
- إجمالي الكميات: ${data.currentTotalQty.toLocaleString("ar-EG")}
- التكلفة الإجمالية الحالية: ${data.currentTotalCost.toLocaleString("ar-EG")} جنيه

🔮 التوقع للعام القادم (بزيادة 20% في الأسعار):
- التكلفة المتوقعة: ${forecast20.toLocaleString("ar-EG")} جنيه
- مقدار الزيادة: ${increase.toLocaleString("ar-EG")} جنيه

🏆 أعلى 5 أصناف تكلفة:
${data.topItems.map((it, i) => `${i + 1}. ${it.name} — ${it.cost.toLocaleString("ar-EG")} ج (كمية: ${it.qty})`).join("\n")}

📅 التوزيع الشهري (أعلى 3 أشهر):
${[...data.monthlyTotals].sort((a, b) => b.total - a.total).slice(0, 3).map((m) => `${m.name}: ${m.total.toLocaleString("ar-EG")}`).join(" | ")}

اكتب تحليلاً موجزاً واحترافياً في 4-6 نقاط قصيرة يتضمن:
1. تقييم الوضع الحالي
2. تفسير التوقع للعام القادم وأثر زيادة الأسعار
3. الأصناف الأكثر تأثيراً على الميزانية
4. توصيات عملية للتحكم في التكاليف والادخار
5. تحذيرات أو فرص محتملة

استخدم رموز تعبيرية بسيطة في بداية كل نقطة. اجعل كل نقطة سطراً واحداً واضحاً.`;

    const { text } = await generateText({
      model: gateway("google/gemini-3-flash-preview"),
      prompt,
    });

    return {
      analysis: text,
      forecast: {
        increase20: forecast20,
        increase15: data.currentTotalCost * 1.15,
        increase10: data.currentTotalCost * 1.1,
        increase5: data.currentTotalCost * 1.05,
        currentCost: data.currentTotalCost,
      },
    };
  });
