using DocumentFormat.OpenXml.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TJADSZY.networkFlow
{
    internal class FordFulkersonDfsSolver : NetworkFlowSolverBase
    {
        public FordFulkersonDfsSolver(int n, int s, int t)
            : base(n, s, t)
        {
        }

        public override void solve()
        {
            for ( double f = dfs(s, INF); f!= 0; f = dfs(s, INF))
            {
                visitedToken++;
                maxFlow = f;
            }
        }

        private double dfs(int node, double flow)
        {
            if (node == t) return flow;

            visited[node] = visitedToken;

            List<Edge> edges = graph[node];
            foreach (Edge edge in edges)
            {
                if (edge.remainingCapacity() > 0 && visited[edge.to] != visitedToken)
                {
                    double bottleNeck = dfs(edge.to, Math.Min(flow, edge.remainingCapacity()));

                    if (bottleNeck > 0)
                    {
                        edge.augment(bottleNeck);
                        return bottleNeck;
                    }
                }
            }
            return 0;
        }
    }
}
