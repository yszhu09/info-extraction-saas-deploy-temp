import { CompareApp } from "@/components/CompareApp";
import { InlineJeProbe } from "@/components/InlineJeProbe";

export const dynamic = "force-dynamic";
export const revalidate = 0;

export default function ComparePage() {
  return (
    <>
      <InlineJeProbe pageName="信息核对" />
      <CompareApp />
    </>
  );
}
